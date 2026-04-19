import argparse
import json
import os
import sys
from datetime import datetime, timezone

from dotenv import load_dotenv

from github_downloader import get_latest_word_reports, get_year_word_reports
from word_comparator import compare_documents
from claude_analyzer import analyze
from report_generator import generate
from email_sender import send_report

load_dotenv()

GH_REPO = os.getenv("GH_REPO", "sinpoussi/mon-repo")
COMPARISON_FILE = "last_comparison.json"


def log(msg):
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    print(f"[{ts}] {msg}")


# ── Sérialisation JSON (datetime → str) ─────────────────────────────────────

def _default(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type non sérialisable : {type(obj)}")


# ── Modes ────────────────────────────────────────────────────────────────────

def cmd_collect():
    log("=== MODE COLLECT ===")

    log(f"Téléchargement des 2 derniers rapports Word depuis '{GH_REPO}'...")
    rapports = get_latest_word_reports(repo_name=GH_REPO, n=2)

    if len(rapports) < 2:
        log(f"ERREUR : {len(rapports)} rapport(s) trouvé(s), 2 requis. Abandon.")
        sys.exit(1)

    doc1, doc2 = rapports[0], rapports[1]
    log(f"Clés disponibles : {list(doc1.keys())}")
    log(f"Document 1 : {doc1['nom']} ({doc1['date_run']})")
    log(f"Document 2 : {doc2['nom']} ({doc2['date_run']})")

    log("Comparaison des documents...")
    diff_data = compare_documents(doc1["contenu_bytes"], doc2["contenu_bytes"])
    log(f"Taux de changement : {diff_data['taux_changement']}%")
    log(f"Ajoutés : {len(diff_data['paragraphes_ajoutes'])} | "
        f"Supprimés : {len(diff_data['paragraphes_supprimes'])} | "
        f"Modifiés : {len(diff_data['paragraphes_modifies'])}")

    log("Analyse Claude en cours...")
    analysis = analyze(doc1["nom"], doc2["nom"], diff_data)
    log(f"Score d'importance : {analysis.get('score_importance')}/10")
    log(f"Alerte : {analysis.get('alerte', 'non').upper()}")

    log("--- Résumé exécutif ---")
    print(analysis.get("resume_executif", ""))
    log("--- Changements importants ---")
    for i, c in enumerate(analysis.get("changements_importants", []), 1):
        print(f"  {i}. {c}")
    log("--- Interprétation business ---")
    print(analysis.get("interpretation", ""))

    payload = {
        "doc1_nom":  doc1["nom"],
        "doc2_nom":  doc2["nom"],
        "diff_data": diff_data,
        "analysis":  analysis,
        "collected_at": datetime.now(timezone.utc).isoformat(),
    }
    with open(COMPARISON_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_default)
    log(f"Résultat sauvegardé dans '{COMPARISON_FILE}'.")


def cmd_rapport_annuel():
    year = datetime.now(timezone.utc).year
    log(f"=== MODE RAPPORT-ANNUEL — {year} ===")

    log(f"Téléchargement de tous les artefacts de l'année {year}...")
    rapports = get_year_word_reports(year=year, repo_name=GH_REPO)

    if len(rapports) < 2:
        log(f"ERREUR : {len(rapports)} rapport(s) trouvé(s) pour {year}, 2 requis. Abandon.")
        sys.exit(1)

    doc1, doc2 = rapports[0], rapports[-1]
    log(f"Premier rapport : {doc1['nom']} ({doc1['date_run']})")
    log(f"Dernier rapport : {doc2['nom']} ({doc2['date_run']})")
    log(f"Nombre total de rapports dans l'année : {len(rapports)}")

    log("Comparaison premier ↔ dernier document...")
    diff_data = compare_documents(doc1["contenu_bytes"], doc2["contenu_bytes"])
    log(f"Taux de changement : {diff_data['taux_changement']}%")

    log("Analyse Claude en cours...")
    analysis = analyze(doc1["nom"], doc2["nom"], diff_data)
    analysis["nb_runs_annee"] = len(rapports)
    analysis["annee"] = year
    analysis["date_premier_run"] = doc1["date_run"].isoformat() if hasattr(doc1["date_run"], "isoformat") else str(doc1["date_run"])
    analysis["date_dernier_run"] = doc2["date_run"].isoformat() if hasattr(doc2["date_run"], "isoformat") else str(doc2["date_run"])

    log(f"Score d'importance : {analysis.get('score_importance')}/10")

    rapport = generate(doc1["nom"], doc2["nom"], diff_data, analysis, "annuel")
    log(f"Objet : {rapport['subject']}")

    log("Envoi de l'email...")
    ok = send_report(rapport["subject"], rapport["body_html"], rapport["body_text"])
    if ok:
        log("Email annuel envoyé avec succès.")
    else:
        log("ERREUR : échec de l'envoi de l'email.")
        sys.exit(1)


def cmd_rapport(type_rapport):
    log(f"=== MODE RAPPORT — {type_rapport.upper()} ===")

    if not os.path.exists(COMPARISON_FILE):
        log(f"ERREUR : '{COMPARISON_FILE}' introuvable. Lancez d'abord 'collect'.")
        sys.exit(1)

    with open(COMPARISON_FILE, "r", encoding="utf-8") as f:
        payload = json.load(f)

    doc1_nom  = payload["doc1_nom"]
    doc2_nom  = payload["doc2_nom"]
    diff_data = payload["diff_data"]
    analysis  = payload["analysis"]
    collected = payload.get("collected_at", "inconnue")

    log(f"Données chargées depuis '{COMPARISON_FILE}' (collecte : {collected})")
    log(f"Génération du rapport {type_rapport}...")

    rapport = generate(doc1_nom, doc2_nom, diff_data, analysis, type_rapport)
    log(f"Objet : {rapport['subject']}")

    log("Envoi de l'email...")
    ok = send_report(rapport["subject"], rapport["body_html"], rapport["body_text"])

    if ok:
        log("Email envoyé avec succès.")
    else:
        log("ERREUR : échec de l'envoi de l'email.")
        sys.exit(1)


# ── Point d'entrée ───────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Agent GitHub Rapports — comparaison et envoi de rapports Word",
    )
    parser.add_argument(
        "mode",
        choices=["collect", "rapport-jour", "rapport-hebdo", "rapport-mensuel", "rapport-annuel"],
        help=(
            "collect : télécharge, compare et analyse les rapports ; "
            "rapport-jour/hebdo/mensuel/annuel : génère et envoie l'email"
        ),
    )
    args = parser.parse_args()

    if args.mode == "collect":
        cmd_collect()
    elif args.mode == "rapport-jour":
        cmd_rapport("quotidien")
    elif args.mode == "rapport-hebdo":
        cmd_rapport("hebdo")
    elif args.mode == "rapport-mensuel":
        cmd_rapport("mensuel")
    elif args.mode == "rapport-annuel":
        cmd_rapport_annuel()


if __name__ == "__main__":
    main()
