import io
import os
import zipfile

import requests
from dotenv import load_dotenv
from github import Github, GithubException

load_dotenv()

GH_TOKEN_PAT = os.getenv("GH_TOKEN_PAT")
GH_REPO = os.getenv("GH_REPO")


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
            run_year = run.created_at.year if run.created_at.tzinfo else run.created_at.year
            if run_year < _year:
                break
            if run_year > _year:
                continue

            try:
                for artifact in run.get_artifacts():
                    zip_url = (
                        f"https://api.github.com/repos/{_repo_name}"
                        f"/actions/artifacts/{artifact.id}/zip"
                    )
                    try:
                        response = requests.get(zip_url, headers=headers, timeout=60)
                        response.raise_for_status()

                        with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
                            for name in zf.namelist():
                                if name.endswith(".docx"):
                                    reports.append({
                                        "nom": os.path.basename(name),
                                        "contenu_bytes": zf.read(name),
                                        "date_run": run.created_at,
                                        "run_number": run.run_number,
                                    })
                                    print(
                                        f"Extrait : {os.path.basename(name)}"
                                        f" (run #{run.run_number}, {run.created_at})"
                                    )
                    except requests.HTTPError as e:
                        print(f"Erreur téléchargement artefact {artifact.id}: {e}")
                    except zipfile.BadZipFile as e:
                        print(f"ZIP invalide pour artefact {artifact.id}: {e}")
                    except Exception as e:
                        print(f"Erreur extraction artefact {artifact.id}: {e}")

            except GithubException as e:
                print(f"Erreur récupération artefacts run #{run.run_number}: {e}")

    except GithubException as e:
        print(f"Erreur GitHub : {e}")
    except Exception as e:
        print(f"Erreur inattendue : {e}")

    reports.sort(key=lambda r: r["date_run"])
    return reports


def get_latest_word_reports(repo_name=None, token=None, n=2):
    """
    Télécharge les artefacts .docx des n derniers workflow runs réussis.

    Retourne une liste de dicts :
        {nom, contenu_bytes, date_run, run_number}
    """
    _token = token or GH_TOKEN_PAT
    _repo_name = repo_name or GH_REPO

    if not _token:
        raise ValueError("GH_TOKEN_PAT manquant (paramètre ou .env)")
    if not _repo_name:
        raise ValueError("GH_REPO manquant (paramètre ou .env)")

    headers = {"Authorization": f"token {_token}"}
    reports = []

    try:
        g = Github(_token)
        repo = g.get_repo(_repo_name)

        runs = repo.get_workflow_runs(status="success")

        recent_runs = []
        for run in runs:
            recent_runs.append(run)
            if len(recent_runs) >= n:
                break

        if not recent_runs:
            print("Aucun workflow run réussi trouvé.")
            return []

        for run in recent_runs:
            try:
                artifacts = run.get_artifacts()
                for artifact in artifacts:
                    zip_url = (
                        f"https://api.github.com/repos/{_repo_name}"
                        f"/actions/artifacts/{artifact.id}/zip"
                    )
                    try:
                        response = requests.get(
                            zip_url, headers=headers, timeout=60
                        )
                        response.raise_for_status()

                        with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
                            for name in zf.namelist():
                                if name.endswith(".docx"):
                                    docx_bytes = zf.read(name)
                                    reports.append({
                                        "nom": os.path.basename(name),
                                        "contenu_bytes": docx_bytes,
                                        "date_run": run.created_at,
                                        "run_number": run.run_number,
                                    })
                                    print(
                                        f"Extrait : {os.path.basename(name)}"
                                        f" (run #{run.run_number},"
                                        f" {run.created_at})"
                                    )
                    except requests.HTTPError as e:
                        print(f"Erreur téléchargement artefact {artifact.id}: {e}")
                    except zipfile.BadZipFile as e:
                        print(f"ZIP invalide pour artefact {artifact.id}: {e}")
                    except Exception as e:
                        print(f"Erreur extraction artefact {artifact.id}: {e}")

            except GithubException as e:
                print(f"Erreur récupération artefacts run #{run.run_number}: {e}")

    except GithubException as e:
        print(f"Erreur GitHub : {e}")
    except Exception as e:
        print(f"Erreur inattendue : {e}")

    return reports
