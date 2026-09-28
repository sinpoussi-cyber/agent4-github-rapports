import io
import os
import zipfile

import requests
from dotenv import load_dotenv
from github import Github, GithubException

load_dotenv()

GH_TOKEN_PAT = os.getenv("GH_TOKEN_PAT")
GH_REPO = os.getenv("GH_REPO")

# Nombre max de runs réussis à parcourir pour trouver `n` rapports .docx.
# Garde-fou : évite de balayer tout l'historique si l'archive est trouée.
# Surchargeable via l'env (utile en backfill).
MAX_RUNS_SCAN = int(os.getenv("GH_MAX_RUNS_SCAN", "80"))


def _extract_docx_reports_from_run(run, headers, repo_name, name_filter=None):
    """Télécharge les artefacts d'un run et renvoie la liste des rapports .docx
    qu'ils contiennent : [{nom, contenu_bytes, date_run, run_number}, ...].

    - Ignore silencieusement les artefacts expirés (téléchargement = 410).
    - `name_filter` : si fourni, ne garde que les .docx dont le nom de fichier
      contient cette sous-chaîne (ex. 'Rapport_Ultimate_BRVM').
    """
    found = []
    try:
        artifacts = run.get_artifacts()
    except GithubException as e:
        print(f"Erreur récupération artefacts run #{run.run_number}: {e}")
        return found

    for artifact in artifacts:
        # Un artefact expiré ne se télécharge plus : on saute sans bruit.
        if getattr(artifact, "expired", False):
            continue

        zip_url = (
            f"https://api.github.com/repos/{repo_name}"
            f"/actions/artifacts/{artifact.id}/zip"
        )
        try:
            response = requests.get(zip_url, headers=headers, timeout=60)
            response.raise_for_status()

            with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
                for name in zf.namelist():
                    if not name.endswith(".docx"):
                        continue
                    base = os.path.basename(name)
                    if name_filter and name_filter not in base:
                        continue
                    found.append({
                        "nom": base,
                        "contenu_bytes": zf.read(name),
                        "date_run": run.created_at,
                        "run_number": run.run_number,
                    })
                    print(
                        f"Extrait : {base}"
                        f" (run #{run.run_number}, {run.created_at})"
                    )
        except requests.HTTPError as e:
            print(f"Erreur téléchargement artefact {artifact.id}: {e}")
        except zipfile.BadZipFile as e:
            print(f"ZIP invalide pour artefact {artifact.id}: {e}")
        except Exception as e:
            print(f"Erreur extraction artefact {artifact.id}: {e}")

    return found


def get_year_word_reports(year=None, repo_name=None, token=None):
    """
    Télécharge les artefacts .docx de TOUS les workflow runs réussis de l'année.

    Retourne une liste triée par date croissante de dicts :
        {nom, contenu_bytes, date_run, run_number}
    """
    from datetime import datetime, timezone as tz

    _token = token or GH_TOKEN_PAT
    _repo_name = repo_name or GH_REPO
    _year = year or datetime.now(tz.utc).year

    if not _token:
        raise ValueError("GH_TOKEN_PAT manquant (paramètre ou .env)")
    if not _repo_name:
        raise ValueError("GH_REPO manquant (paramètre ou .env)")

    headers = {"Authorization": f"token {_token}"}
    reports = []

    try:
        g = Github(_token)
        repo = g.get_repo(_repo_name)

        for run in repo.get_workflow_runs(status="success"):
            run_year = run.created_at.year
            if run_year < _year:
                break
            if run_year > _year:
                continue
            reports.extend(
                _extract_docx_reports_from_run(run, headers, _repo_name)
            )

    except GithubException as e:
        print(f"Erreur GitHub : {e}")
    except Exception as e:
        print(f"Erreur inattendue : {e}")

    reports.sort(key=lambda r: r["date_run"])
    return reports


# Préfixe des artefacts publiés par brvm-analysis-suite
# (step « Upload Rapport » : name: rapport-brvm-${{ github.run_number }}).
ARTIFACT_PREFIX = os.getenv("GH_ARTIFACT_PREFIX", "rapport-brvm-")
MAX_ARTIFACTS_SCAN = int(os.getenv("GH_MAX_ARTIFACTS_SCAN", "100"))


def _parse_gh_date(value):
    from datetime import datetime
    return datetime.fromisoformat(str(value).replace("Z", "+00:00"))


def _latest_reports_via_artifacts(repo_name, token, n, name_filter=None):
    """Sélection par l'API des ARTEFACTS, triés par date de création réelle.

    Pourquoi : l'ancienne méthode prenait les premiers runs « success » dans
    l'ordre renvoyé par l'API des workflow runs, sans contrôle de date. Le
    28/09/2026, elle a ainsi renvoyé les runs #413/#412 (5 et 4 sept.) au lieu
    du run #438 du jour. Ici on ne dépend ni de cet ordre ni du statut du run :
    un artefact rapport-brvm-* n'existe que si le rapport Word a été produit.
    """
    headers = {"Authorization": f"token {token}",
               "Accept": "application/vnd.github+json"}
    url = f"https://api.github.com/repos/{repo_name}/actions/artifacts"
    candidats, page = [], 1
    while len(candidats) < MAX_ARTIFACTS_SCAN and page <= 5:
        r = requests.get(url, headers=headers, timeout=30,
                         params={"per_page": 100, "page": page})
        r.raise_for_status()
        items = r.json().get("artifacts", [])
        if not items:
            break
        for a in items:
            if a.get("expired") or not str(a.get("name", "")).startswith(ARTIFACT_PREFIX):
                continue
            candidats.append(a)
        page += 1

    # Du plus récent au plus ancien, selon la date réelle de l'artefact
    candidats.sort(key=lambda a: _parse_gh_date(a["created_at"]), reverse=True)

    reports = []
    for a in candidats:
        if len(reports) >= n:
            break
        try:
            resp = requests.get(a["archive_download_url"], headers=headers, timeout=60)
            resp.raise_for_status()
            suffixe = a["name"][len(ARTIFACT_PREFIX):]
            run_number = int(suffixe) if suffixe.isdigit() else None
            date_run = _parse_gh_date(a["created_at"])
            with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
                for name in zf.namelist():
                    base = os.path.basename(name)
                    if not name.endswith(".docx") or (name_filter and name_filter not in base):
                        continue
                    reports.append({"nom": base, "contenu_bytes": zf.read(name),
                                    "date_run": date_run, "run_number": run_number})
                    print(f"Extrait : {base} (artefact {a['name']}, {date_run})")
        except Exception as e:
            print(f"Erreur artefact {a.get('name')}: {e}")

    reports.sort(key=lambda r: r["date_run"], reverse=True)
    return reports[:n]


def get_latest_word_reports(repo_name=None, token=None, n=2,
                            name_filter=None, max_runs_scan=None):
    """
    Télécharge les artefacts .docx des workflow runs réussis, en parcourant
    l'historique du plus récent au plus ancien jusqu'à réunir `n` RAPPORTS
    (et non `n` runs). Un run réussi sans .docx (autre workflow du dépôt,
    artefact expiré, run interne sans rapport) est simplement ignoré au lieu
    de faire échouer la collecte.

    Retourne une liste de `n` dicts max, du plus récent au plus ancien :
        {nom, contenu_bytes, date_run, run_number}
    """
    _token = token or GH_TOKEN_PAT
    _repo_name = repo_name or GH_REPO
    _cap = max_runs_scan or MAX_RUNS_SCAN

    if not _token:
        raise ValueError("GH_TOKEN_PAT manquant (paramètre ou .env)")
    if not _repo_name:
        raise ValueError("GH_REPO manquant (paramètre ou .env)")

    # 1) Méthode principale : artefacts triés par date réelle
    try:
        reports = _latest_reports_via_artifacts(_repo_name, _token, n, name_filter)
        if len(reports) >= n:
            print(f"{len(reports)} rapport(s) le(s) plus récent(s) via l'API des artefacts.")
            return reports
        print(f"API artefacts : {len(reports)} rapport(s) seulement — repli sur le parcours des runs.")
    except Exception as e:
        print(f"API artefacts indisponible ({e}) — repli sur le parcours des runs.")

    # 2) Repli : ancienne méthode (runs « success »)
    headers = {"Authorization": f"token {_token}"}
    reports = []

    try:
        g = Github(_token)
        repo = g.get_repo(_repo_name)

        runs_scanned = 0
        runs_with_report = 0
        for run in repo.get_workflow_runs(status="success"):
            if len(reports) >= n or runs_scanned >= _cap:
                break
            runs_scanned += 1

            run_reports = _extract_docx_reports_from_run(
                run, headers, _repo_name, name_filter=name_filter)
            if run_reports:
                runs_with_report += 1
                # Un run = un rapport journalier : on prend le plus pertinent
                # (unique en pratique) et on complète jusqu'à `n`.
                reports.extend(run_reports)

        if not reports:
            print("Aucun rapport .docx trouvé dans les runs réussis récents.")
        else:
            print(
                f"{len(reports)} rapport(s) réuni(s) sur {runs_with_report} "
                f"run(s) porteur(s), {runs_scanned} run(s) parcouru(s)."
            )

    except GithubException as e:
        print(f"Erreur GitHub : {e}")
    except Exception as e:
        print(f"Erreur inattendue : {e}")

    # Ordre décroissant garanti (le plus récent en premier) puis on tronque.
    reports.sort(key=lambda r: r["date_run"], reverse=True)
    return reports[:n]
