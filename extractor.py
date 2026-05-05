"""
Couche d'extraction LLM.
Responsabilité unique : appeler Claude et retourner un JSON minimal par société.
Champs retournés : ticker, nom, secteur, cours, var_1j, reco, score.
"""
import json
import os
import re

import anthropic
from dotenv import load_dotenv

load_dotenv()

_MODEL = "claude-sonnet-4-20250514"
_BATCH_SIZE = 3
_MAX_RETRIES = 3


# ── Utilitaires JSON ──────────────────────────────────────────────────────────

def clean_json_string(raw: str) -> str:
    """Nettoie une réponse LLM pour isoler un tableau JSON valide."""
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    # "nul" seul → "null"  (évite de toucher à "null" existant ou des mots comme "nulité")
    raw = re.sub(r'(?<!["\w])nul(?![\w"l])', "null", raw)
    start = raw.find("[")
    end = raw.rfind("]") + 1
    if start == -1 or end == 0:
        return raw
    return raw[start:end]


def safe_json_load(text: str) -> list:
    """Parse JSON en retournant [] en cas d'erreur de parsing."""
    try:
        result = json.loads(text)
        return result if isinstance(result, list) else []
    except (json.JSONDecodeError, ValueError) as e:
        print(f"  [Extractor/JSON] Erreur parsing : {e}")
        return []


# ── Extraction des tickers ────────────────────────────────────────────────────

def get_tickers(full_text: str) -> list:
    """Appelle le LLM pour extraire la liste des tickers BRVM présents dans le texte."""
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    msg = client.messages.create(
        model=_MODEL,
        max_tokens=512,
        messages=[{
            "role": "user",
            "content": (
                "Extrais UNIQUEMENT la liste des tickers/symboles boursiers de toutes "
                "les sociétés BRVM présentes dans ce rapport.\n"
                "Réponds UNIQUEMENT avec du JSON valide. Ne mets aucun texte hors JSON.\n"
                "Utilise null et jamais nul pour les valeurs absentes.\n"
                'Retourne UNIQUEMENT ce JSON : {"tickers": ["SGBCI", "SONATEL", ...]}\n\n'
                f"RAPPORT :\n{full_text[:20000]}"
            ),
        }],
    )
    raw = msg.content[0].text.strip()
    print(f"  [Extractor/Tickers] Réponse ({len(raw)} chars) : {raw[:200]}")
    cleaned = re.sub(r"```json\s*", "", raw)
    cleaned = re.sub(r"```\s*", "", cleaned).strip()
    start, end = cleaned.find("{"), cleaned.rfind("}") + 1
    if start == -1 or end == 0:
        print("  [Extractor/Tickers] ERREUR : aucun objet JSON trouvé")
        return []
    try:
        tickers = json.loads(cleaned[start:end]).get("tickers", [])
        print(f"  [Extractor/Tickers] {len(tickers)} ticker(s) : {tickers[:10]}")
        return tickers
    except json.JSONDecodeError as e:
        print(f"  [Extractor/Tickers] JSONDecodeError : {e}")
        return []


# ── Extraction par batch ──────────────────────────────────────────────────────

def extract_batch(full_text: str, tickers: list, freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Extrait les données minimales pour un batch de sociétés (max _BATCH_SIZE).
    Retourne une liste de dicts {ticker, nom, secteur, cours, var_1j, reco, score}.
    Retry jusqu'à _MAX_RETRIES fois. Fallback minimal si tout échoue.
    """
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    period_ctx = ""
    if freq != "JOUR":
        _descs = {
            "HEBDO": "7 derniers jours",
            "MENSUEL": "30 derniers jours",
            "TRIM": "dernier trimestre",
            "ANNUEL": "dernière année",
        }
        nb = (period_info or {}).get("nb_seances", "?")
        period_ctx = f"Période analysée : {_descs.get(freq, freq)} ({nb} séances).\n"

    prompt = (
        f"{period_ctx}"
        f"Rapport BRVM — Extrais les données des sociétés suivantes : {', '.join(tickers)}\n\n"
        f"RAPPORT :\n{full_text[:35000]}\n\n"
        "CONTRAINTES STRICTES :\n"
        "- Réponds UNIQUEMENT en JSON valide.\n"
        "- Pas de texte, commentaire ou explication hors du JSON.\n"
        "- Utilise null (JAMAIS nul) pour toutes les valeurs absentes.\n"
        "- La réponse doit commencer par [ et terminer par ].\n\n"
        "Schéma attendu — retourne UNIQUEMENT ce tableau JSON :\n"
        "[\n"
        '  {"ticker": "SGBCI", "nom": "Société Générale de Banques", "secteur": "Banque",\n'
        '   "cours": "14500", "var_1j": "+0.5%", "reco": "ACHAT", "score": 82},\n'
        '  {"ticker": "SONATEL", "nom": "Sonatel", "secteur": "Télécommunications",\n'
        '   "cours": "18000", "var_1j": "-0.3%", "reco": "NEUTRE", "score": 65}\n'
        "]\n"
    )

    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            msg = client.messages.create(
                model=_MODEL,
                max_tokens=512,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            print(f"  [Extractor/Batch] Tentative {attempt}/{_MAX_RETRIES} — {len(raw)} chars")

            cleaned = clean_json_string(raw)
            print(f"  [Extractor/Batch] JSON nettoyé ({len(cleaned)} chars) : {cleaned[:150]}")

            parsed = safe_json_load(cleaned)
            if parsed:
                print(f"  [Extractor/Batch] OK : {len(parsed)} société(s)")
                return parsed

            print(f"  [Extractor/Batch] Résultat vide — tentative {attempt}/{_MAX_RETRIES}")
        except Exception as e:
            print(f"  [Extractor/Batch] Erreur tentative {attempt}/{_MAX_RETRIES} : {e}")

    print(f"  [Extractor/Batch] Toutes tentatives échouées — fallback pour : {tickers}")
    return [{"ticker": t, "nom": t, "secteur": None, "cours": None, "var_1j": None,
             "reco": "INCONNU", "score": 0} for t in tickers]


# ── Pipeline complet d'extraction ────────────────────────────────────────────

def extract_all(full_text: str, freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Pipeline : get_tickers() → extract_batch() par tranches de _BATCH_SIZE.
    Retourne la liste brute des sociétés (JSON minimal).
    """
    print("  [Extractor] Étape A : Récupération des tickers...")
    tickers = get_tickers(full_text)
    if not tickers:
        print("  [Extractor] AVERTISSEMENT : aucun ticker trouvé.")
        return []
    print(f"  [Extractor] {len(tickers)} ticker(s) identifié(s).")

    all_companies = []
    total_batches = (len(tickers) + _BATCH_SIZE - 1) // _BATCH_SIZE
    for i in range(0, len(tickers), _BATCH_SIZE):
        batch = tickers[i:i + _BATCH_SIZE]
        batch_num = i // _BATCH_SIZE + 1
        print(f"  [Extractor] Étape B — Batch {batch_num}/{total_batches} : {', '.join(batch)}")
        companies = extract_batch(full_text, batch, freq, period_info)
        print(f"  [Extractor] Batch {batch_num} → {len(companies)} société(s) extraite(s)")
        all_companies.extend(companies)

    print(f"  [Extractor] Total LLM : {len(all_companies)} société(s) extraite(s).")
    return all_companies
