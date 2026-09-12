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
