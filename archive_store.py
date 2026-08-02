"""
archive_store.py — Archivage durable des rapports BRVM dans le dépôt GitHub
===========================================================================
Les artefacts GitHub Actions expirent (90 jours maximum). Pour construire des
rapports MENSUELS et ANNUELS fiables, on archive les documents directement dans
un dépôt GitHub (commits via l'API), rangés par date.

Arborescence dans le dépôt d'archive :
  archives/daily/AAAA/MM/AAAA-MM-JJ_<nom>.docx      ← rapports journaliers sources
  archives/monthly/AAAA/AAAA-MM_Note_Mensuelle.docx ← notes stratégiques mensuelles

Variables d'environnement :
  GH_TOKEN_PAT   — token avec droit d'écriture sur le dépôt (contents: write)
  ARCHIVE_REPO   — dépôt d'archive "owner/repo"
                   (défaut : GITHUB_REPOSITORY, i.e. le dépôt courant de l'Action)
  ARCHIVE_BRANCH — branche cible (défaut : branche par défaut du dépôt)

Toutes les fonctions importent PyGithub à l'intérieur du corps : le module
s'importe donc sans dépendance, et un échec d'archivage ne doit jamais faire
échouer l'envoi du rapport quotidien (l'appelant encapsule dans try/except).
"""

import base64
import os
from datetime import date, datetime

DAILY_PREFIX = "archives/daily"
MONTHLY_PREFIX = "archives/monthly"


# ── Accès dépôt ───────────────────────────────────────────────────────────────

def _repo():
    from github import Github
    token = os.getenv("GH_TOKEN_PAT")
    if not token:
        raise ValueError("GH_TOKEN_PAT manquant — archivage impossible.")
    repo_name = os.getenv("ARCHIVE_REPO") or os.getenv("GITHUB_REPOSITORY")
    if not repo_name:
        raise ValueError("ARCHIVE_REPO / GITHUB_REPOSITORY manquant — archivage impossible.")
    return Github(token).get_repo(repo_name)


def _branch(repo):
    return os.getenv("ARCHIVE_BRANCH") or repo.default_branch


def _content_bytes(repo, content_file):
    """Retourne les octets d'un ContentFile, avec repli via l'API git blob
    pour les fichiers volumineux dont le contenu n'est pas inline."""
    try:
        data = content_file.decoded_content
        if data:
            return data
    except Exception:
        pass
    blob = repo.get_git_blob(content_file.sha)
    return base64.b64decode(blob.content)


def _commit_file(path, content_bytes, message):
    """Crée ou met à jour un fichier binaire dans le dépôt (idempotent)."""
    from github import GithubException
    repo = _repo()
    branch = _branch(repo)
    try:
        existing = repo.get_contents(path, ref=branch)
        repo.update_file(path, message, content_bytes, existing.sha, branch=branch)
        return "updated"
    except GithubException as e:
        if getattr(e, "status", None) == 404:
            repo.create_file(path, message, content_bytes, branch=branch)
            return "created"
        raise


# ── Helpers date ──────────────────────────────────────────────────────────────

def _as_date(d):
    if isinstance(d, datetime):
        return d.date()
    if isinstance(d, date):
        return d
    return date.today()


def _date_from_name(name, year, month):
    """Nom attendu : 'AAAA-MM-JJ_...'. Repli sur le 1er du mois si non parsable."""
    try:
        return date.fromisoformat(name[:10])
    except (ValueError, TypeError):
        return date(year, month, 1)


# ── Écriture ──────────────────────────────────────────────────────────────────

def archive_daily_report(nom, contenu_bytes, date_run=None):
    """Archive un rapport journalier source. Retourne (path, action)."""
    d = _as_date(date_run)
    safe = os.path.basename(nom or "rapport.docx")
    path = f"{DAILY_PREFIX}/{d.year:04d}/{d.month:02d}/{d.isoformat()}_{safe}"
    action = _commit_file(path, contenu_bytes, f"archive: rapport journalier {d.isoformat()}")
    return path, action


def archive_monthly_note(contenu_bytes, year, month, nom=None):
    """Archive une note stratégique mensuelle. Retourne (path, action)."""
    fname = os.path.basename(nom) if nom else f"{year:04d}-{month:02d}_Note_Mensuelle.docx"
    path = f"{MONTHLY_PREFIX}/{year:04d}/{fname}"
    action = _commit_file(path, contenu_bytes, f"archive: note mensuelle {year:04d}-{month:02d}")
    return path, action


# ── Lecture ───────────────────────────────────────────────────────────────────

def _list_docx(folder, sort=True):
    """Retourne la liste des ContentFile .docx d'un dossier ([] si absent)."""
    from github import GithubException
    repo = _repo()
    branch = _branch(repo)
    try:
        items = repo.get_contents(folder, ref=branch)
    except GithubException as e:
        if getattr(e, "status", None) == 404:
            return repo, []
        raise
    if isinstance(items, list):
        docx = [it for it in items if it.name.endswith(".docx")]
    else:  # dossier réduit à un seul fichier
        docx = [items] if items.name.endswith(".docx") else []
    if sort:
        docx.sort(key=lambda c: c.name)
    return repo, docx


def get_daily_reports_for_month(year, month):
    """Retourne [{nom, contenu_bytes, date_run}] (ordre chronologique CROISSANT)
    pour tous les rapports journaliers archivés du mois civil demandé."""
    folder = f"{DAILY_PREFIX}/{year:04d}/{month:02d}"
    repo, docx = _list_docx(folder)
    out = []
    for it in docx:
        out.append({
            "nom": it.name,
            "contenu_bytes": _content_bytes(repo, it),
            "date_run": _date_from_name(it.name, year, month),
        })
    out.sort(key=lambda r: r["date_run"])
    return out


def get_monthly_notes_for_year(year):
    """Retourne [{nom, contenu_bytes}] (ordre chronologique CROISSANT) pour toutes
    les notes stratégiques mensuelles archivées de l'année demandée."""
    folder = f"{MONTHLY_PREFIX}/{year:04d}"
    repo, docx = _list_docx(folder)
    return [{"nom": it.name, "contenu_bytes": _content_bytes(repo, it)} for it in docx]
