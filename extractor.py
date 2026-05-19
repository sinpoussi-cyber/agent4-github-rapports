"""
Couche d'extraction LLM.
Pipeline : nettoyage -> segmentation -> extraction JSON ciblée par société.
Champs : ticker, nom, secteur, cours, var_1j, reco, score, mm, boll, macd, rsi, stoch.
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

# Seuils validation qualité
_MAX_NULL_RATIO = 0.2   # 80% des champs essentiels doivent être présents
_MIN_SCORE_RATIO = 0.3

# Limites de taille par appel LLM
_MAX_CHARS_TICKERS   = 25_000
_MAX_CHARS_BRVM_GLOB = 15_000
_MAX_CHARS_BATCH     = 35_000
_CHARS_PER_TICKER    =  6_000  # fenêtre fallback (occurrences génériques)
_SECTION_WINDOW      = 10_000  # fenêtre avant -> couvre la section dédiée


# ═══════════════════════════════════════════════════════════════════════════════
# NETTOYAGE DU TEXTE
# ═══════════════════════════════════════════════════════════════════════════════

_EMOJI_RE = re.compile(
    "[\U0001F300-\U0001F5FF\U0001F600-\U0001F64F\U0001F680-\U0001F6FF"
    "\U0001F700-\U0001F77F\U0001F780-\U0001F7FF\U0001F800-\U0001F8FF"
    "\U0001F900-\U0001F9FF\U0001FA00-\U0001FA6F\U0001FA70-\U0001FAFF"
    "\U00002702-\U000027B0\U000024C2-\U0001F251]+",
    flags=re.UNICODE,
)
_DECO_LINE_RE  = re.compile(r'^[ \t]*[=\-\*_#~\|]{4,}[ \t]*$', re.MULTILINE)
_PAGE_NO_RE    = re.compile(r'^\s*\d{1,3}\s*$', re.MULTILINE)
_MULTI_SPC_RE  = re.compile(r'[ \t]{2,}')
_MULTI_NL_RE   = re.compile(r'\n{3,}')


def clean_text(text: str) -> str:
    """
    Supprime :
    - emojis et symboles Unicode décoratifs
    - lignes de séparation pure (====, ----)
    - numéros de page isolés
    - espaces multiples
    - lignes consécutives identiques (headers répétés)
    """
    size_before = len(text)

    text = _EMOJI_RE.sub(" ", text)
    text = _DECO_LINE_RE.sub("", text)
    text = _PAGE_NO_RE.sub("", text)
    text = _MULTI_SPC_RE.sub(" ", text)
    text = _MULTI_NL_RE.sub("\n\n", text)

    # Dédupliquer les lignes consécutives identiques
    lines, deduped, prev = text.split('\n'), [], None
    for line in lines:
        stripped = line.strip()
        if stripped and stripped == prev:
            continue
        deduped.append(line)
        if stripped:
            prev = stripped

    text = '\n'.join(deduped).strip()
    reduction = (1 - len(text) / max(size_before, 1)) * 100
    print(f"  [TextClean] {size_before:,} -> {len(text):,} chars ({reduction:.1f}% reduit)")
    return text


# ═══════════════════════════════════════════════════════════════════════════════
# SEGMENTATION EN SECTIONS
# ═══════════════════════════════════════════════════════════════════════════════

_COMPANY_SECTION_RE = re.compile(
    r'(analyse\s+(individuelle|par\s+soci[eé]t[eé]|des\s+valeurs)|'
    r'fiche\s+(individuelle|par\s+valeur|soci[eé]t[eé])|'
    r'recommandations?\s+d[eé]taill[eé]|'
    r'top\s+(opportunit|valeur|titre)|'
    r'soci[eé]t[eé]s?\s+analys|'
    r'analyse\s+technique\s+et\s+fondamentale)',
    re.IGNORECASE,
)
_NEWS_SECTION_RE = re.compile(
    r'(actualit[eé]s?\s*(du\s+march[eé])?|communiqu[eé]s?\s+de\s+presse|'
    r'news\s+du\s+march[eé]|informations?\s+boursi[eè]res?)',
    re.IGNORECASE,
)
# Ligne contenant un ticker uppercase + un prix ou une recommandation
_TICKER_DATA_RE = re.compile(
    r'\b[A-Z]{2,8}\b.{0,60}?(\d{3,}|ACHAT|VENTE|NEUTRE)',
    re.IGNORECASE,
)


def split_text_sections(text: str) -> dict:
    """
    Segmente le texte nettoyé en trois sections logiques.

    Stratégie de détection (ordre de priorité) :
    1. Marqueurs de section explicites (regex)
    2. Première ligne ticker + données financières (après 10 % du texte)
    3. Fallback positionnel : 25 % globale / 75 % sociétés

    Retourne :
    {
        "brvm_global"  : str,   # indice, capitalisation, données macro
        "societes"     : str,   # données par entreprise
        "actualites"   : str,   # news (peut être vide)
        "full_clean"   : str,   # texte complet nettoyé
        "stats"        : dict,
    }
    """
    lines = text.split('\n')
    n = len(lines)

    brvm_end      = None
    company_start = None
    news_start    = None

    for i, line in enumerate(lines):
        if news_start is None and _NEWS_SECTION_RE.search(line):
            news_start = i

        if company_start is None and _COMPANY_SECTION_RE.search(line):
            company_start = i
            brvm_end = brvm_end or i

        # Transition ticker+données détectée après les 10 premières %
        if company_start is None and i > n * 0.10 and _TICKER_DATA_RE.search(line):
            company_start = i
            brvm_end = brvm_end or i

    # Fallback positionnel
    if company_start is None:
        brvm_end      = max(1, n // 4)
        company_start = brvm_end

    # Garantir une section globale d'au moins 50 lignes ou 3000 chars
    _MIN_BRVM_LINES = 50
    _MIN_BRVM_CHARS = 3_000
    brvm_end = max(brvm_end or 0, min(_MIN_BRVM_LINES, n))
    while brvm_end < n and len('\n'.join(lines[:brvm_end])) < _MIN_BRVM_CHARS:
        brvm_end = min(brvm_end + 10, n)

    brvm_section    = '\n'.join(lines[:brvm_end]).strip()
    company_section = '\n'.join(
        lines[company_start: (news_start if news_start else n)]
    ).strip()
    news_section    = '\n'.join(lines[news_start:]).strip() if news_start else ""

    stats = {
        "total_chars"       : len(text),
        "brvm_global_chars" : len(brvm_section),
        "societes_chars"    : len(company_section),
        "actualites_chars"  : len(news_section),
        "total_lines"       : n,
        "brvm_end_line"     : brvm_end,
        "company_start_line": company_start,
        "news_start_line"   : news_start,
    }

    print(f"  [TextSplit] {n} lignes, {len(text):,} chars au total — "
          f"{len([s for s in [brvm_section, company_section, news_section] if s])} section(s) détectée(s)")
    print(f"  [TextSplit] BRVM global : {stats['brvm_global_chars']:,} chars "
          f"(lignes 0 -> {brvm_end})")
    print(f"  [TextSplit] Sociétés   : {stats['societes_chars']:,} chars "
          f"(lignes {company_start} -> {news_start or n})")
    if news_section:
        print(f"  [TextSplit] Actualités : {stats['actualites_chars']:,} chars")

    return {
        "brvm_global" : brvm_section,
        "societes"    : company_section,
        "actualites"  : news_section,
        "full_clean"  : text,
        "stats"       : stats,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION CONTEXTE CIBLÉ PAR TICKER
# ═══════════════════════════════════════════════════════════════════════════════

def build_ticker_context(sections: dict, tickers: list) -> str:
    """
    Pour chaque ticker, assemble un contexte ciblé en trois sources :
    1. Ligne résumé du top dans brvm_global (prix, reco, confiance, risque)
    2. Section dédiée à la société (en-tête numéroté "N. TICKER -") — fenêtre
       _SECTION_WINDOW chars vers l'avant ; on prend le DERNIER match pour
       éviter la table des matières.
    3. Fallback : occurrences génériques avec fenêtre _CHARS_PER_TICKER.

    Retourne un texte ciblé < _MAX_CHARS_BATCH chars.
    """
    full_text     = sections.get("full_clean", "")
    brvm_text     = sections.get("brvm_global") or ""
    societes_text = sections.get("societes") or ""
    contexts      = []
    half          = _CHARS_PER_TICKER // 2

    for ticker in tickers:
        snippets = []
        sources_used = []

        # 1. Ligne résumé top (Prix / ACHAT / VENTE / NEUTRE)
        summary_pat = re.compile(
            r'^[^\n]*\b' + re.escape(ticker) + r'\b[^\n]*'
            r'(?:\d{2,}\s*FCFA|ACHAT|VENTE|NEUTRE)[^\n]*$',
            re.MULTILINE | re.IGNORECASE,
        )
        m_sum = summary_pat.search(brvm_text or full_text)
        if m_sum:
            snippets.append(f"[Résumé top] {m_sum.group(0).strip()}")
            sources_used.append("résumé")

        # 2. Section dédiée — dernier match pour éviter la TOC
        header_pat = re.compile(
            r'(?:^|\n)\s*\d+\.\s+' + re.escape(ticker) + r'\s*[-—]',
            re.MULTILINE,
        )
        header_matches = list(header_pat.finditer(full_text))
        if header_matches:
            start = header_matches[-1].start()
            end   = min(len(full_text), start + _SECTION_WINDOW)
            snippets.append(full_text[start:end])
            sources_used.append("section")

        # 3. Section "Analyse financière" dédiée (pages 42-66 du rapport source)
        fin_pat = re.compile(
            r'\banalyse\s+financi[èe]re\b[^\n]{0,400}\b' + re.escape(ticker) + r'\b'
            r'|'
            r'\b' + re.escape(ticker) + r'\b[^\n]{0,400}\banalyse\s+financi[èe]re\b',
            re.IGNORECASE,
        )
        fin_match = fin_pat.search(full_text)
        if fin_match:
            start = max(0, fin_match.start() - 500)
            end   = min(len(full_text), fin_match.end() + 4500)
            snippets.append(f"[Analyse financière] {full_text[start:end]}")
            sources_used.append("financière")

        # 4. Fallback générique si rien trouvé
        if not snippets:
            pat = re.compile(r'\b' + re.escape(ticker) + r'\b')
            source = societes_text or full_text
            for m in list(pat.finditer(source))[:2]:
                a = max(0, m.start() - half)
                b = min(len(source), m.end() + half)
                snippets.append(source[a:b])
            if snippets:
                sources_used.append("fallback")

        if snippets:
            ctx = f"=== {ticker} ({', '.join(sources_used)}) ===\n" + "\n\n---\n\n".join(snippets)
        else:
            print(f"  [TickerCtx] '{ticker}' introuvable dans le texte")
            ctx = f"=== {ticker} : non trouvé dans le rapport ==="

        contexts.append(ctx)

    combined = "\n\n".join(contexts)[:_MAX_CHARS_BATCH]
    print(f"  [TickerCtx] Contexte pour {tickers} : {len(combined):,} chars")
    return combined


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES JSON
# ═══════════════════════════════════════════════════════════════════════════════

def clean_json_string(raw: str) -> str:
    """Nettoie la réponse LLM pour isoler un tableau JSON valide."""
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    raw = re.sub(r'(?<!["\w])nul(?![\w"l])', "null", raw)
    start, end = raw.find("["), raw.rfind("]") + 1
    if start == -1 or end == 0:
        return raw
    return raw[start:end]


def safe_json_load(text: str) -> list:
    try:
        result = json.loads(text)
        return result if isinstance(result, list) else []
    except (json.JSONDecodeError, ValueError) as e:
        print(f"  [Extractor/JSON] Erreur parsing : {e}")
        print(f"  [Extractor/JSON] Texte reçu : {text[:200]}")
        return []


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION LLM
# ═══════════════════════════════════════════════════════════════════════════════

def get_tickers(sections: dict) -> list:
    """
    Identifie les tickers BRVM depuis la section sociétés.
    Utilise au plus _MAX_CHARS_TICKERS chars pour couvrir tous les tickers
    même dans un rapport long.
    """
    text = sections.get("societes") or sections.get("full_clean", "")
    snippet = text[:_MAX_CHARS_TICKERS]

    print(f"  [Extractor/Tickers] Texte utilisé : {len(snippet):,} chars")
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    msg = client.messages.create(
        model=_MODEL,
        max_tokens=512,
        messages=[{
            "role": "user",
            "content": (
                "Extrais UNIQUEMENT la liste des tickers/symboles boursiers de toutes "
                "les sociétés BRVM présentes dans ce rapport.\n"
                "Réponds UNIQUEMENT avec du JSON valide. Aucun texte hors JSON.\n"
                'Retourne UNIQUEMENT : {"tickers": ["SGBCI", "SONATEL", ...]}\n\n'
                f"RAPPORT :\n{snippet}"
            ),
        }],
    )
    raw = msg.content[0].text.strip()
    print(f"  [Extractor/Tickers] Réponse ({len(raw)} chars) : {raw[:200]}")
    cleaned = re.sub(r"```(?:json)?\s*", "", raw).strip()
    s, e = cleaned.find("{"), cleaned.rfind("}") + 1
    if s == -1 or e == 0:
        print("  [Extractor/Tickers] ERREUR : aucun objet JSON")
        return []
    try:
        tickers = json.loads(cleaned[s:e]).get("tickers", [])
        print(f"  [Extractor/Tickers] {len(tickers)} ticker(s) : {tickers[:15]}")
        return tickers
    except json.JSONDecodeError as e2:
        print(f"  [Extractor/Tickers] JSONDecodeError : {e2}")
        return []


def extract_brvm_global(sections: dict) -> dict:
    """
    Extrait les données de marché globales BRVM.
    Utilise uniquement la section brvm_global (courte et ciblée).
    """
    brvm_text = sections.get("brvm_global") or sections.get("full_clean", "")
    snippet   = brvm_text[:_MAX_CHARS_BRVM_GLOB]
    print(f"  [Extractor/BRVM] Section globale : {len(brvm_text):,} chars -> envoi {len(snippet):,} chars")

    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    prompt = (
        "Tu analyses un rapport boursier BRVM. Extrais les données globales du marché.\n\n"
        "RÈGLES STRICTES :\n"
        "- Réponds UNIQUEMENT en JSON valide (objet {}).\n"
        "- Utilise null pour toute valeur absente — NE PAS inventer.\n\n"
        "Schéma attendu :\n"
        "{\n"
        '  "brvm_composite"  : "valeur ou variation du BRVM Composite",\n'
        '  "perf_100j"       : "performance sur 100 jours ou période disponible",\n'
        '  "signaux_achat"   : <entier ou null>,\n'
        '  "signaux_vente"   : <entier ou null>,\n'
        '  "signaux_neutre"  : <entier ou null>,\n'
        '  "top_opportunite" : "ticker ou nom de la meilleure opportunité",\n'
        '  "secteur_leader"  : "secteur le plus performant"\n'
        "}\n\n"
        f"TEXTE SOURCE :\n{snippet}"
    )
    try:
        msg = client.messages.create(
            model=_MODEL, max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        raw     = msg.content[0].text.strip()
        cleaned = re.sub(r"```(?:json)?\s*", "", raw).strip()
        s, e    = cleaned.find("{"), cleaned.rfind("}") + 1
        if s != -1 and e > 0:
            data = json.loads(cleaned[s:e])
            print(f"  [Extractor/BRVM] Données : {data}")
            return data
    except Exception as exc:
        print(f"  [Extractor/BRVM] Erreur : {exc}")
    return {}


def extract_batch(sections: dict, tickers: list,
                  freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Extrait les données complètes pour un batch de tickers.
    Utilise le contexte ciblé (build_ticker_context) plutôt que le texte brut complet.
    max_tokens = 2048 pour accueillir toutes les données sans troncature.
    """
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    period_ctx = ""
    if freq != "JOUR":
        _descs = {"HEBDO": "7 derniers jours", "MENSUEL": "30 derniers jours",
                  "TRIM": "dernier trimestre", "ANNUEL": "dernière année"}
        nb = (period_info or {}).get("nb_seances", "?")
        period_ctx = f"Période : {_descs.get(freq, freq)} ({nb} séances).\n"

    ticker_context = build_ticker_context(sections, tickers)
    tickers_str    = ", ".join(tickers)

    prompt = (
        f"{period_ctx}"
        f"Tu analyses un rapport boursier BRVM. "
        f"Extrais les données RÉELLES des sociétés : {tickers_str}\n\n"
        "RÈGLES STRICTES :\n"
        "1. Réponds UNIQUEMENT en JSON. Commence par [ et termine par ].\n"
        "2. null (JAMAIS 0 par défaut) pour toute valeur absente dans le texte.\n"
        "3. 'score' est OBLIGATOIRE et ne peut pas être 0 :\n"
        "   ACHAT -> >= 60 | NEUTRE -> 40-59 | VENTE -> <= 39\n"
        "   Si un score explicite est dans le texte, utilise-le.\n"
        "4. 'cours' : cours actuel en FCFA (nombre, sans unité).\n"
        "5. 'var_1j' : variation du cours sur la DERNIÈRE SÉANCE uniquement "
        "(format avec signe : '+0.29%', '-1.2%'). JAMAIS plus de ±10% en une "
        "journée — si la valeur lue dépasse +10% ou -10%, c'est une performance "
        "multi-jours (ex: perf_100j) mal interprétée : retourner null.\n"
        "6. mm/boll/macd/rsi/stoch : 'haussier', 'neutre' ou 'baissier'. "
        "null si non mentionné.\n\n"
        "Schéma exact :\n"
        '[\n  {"ticker":"SGBCI","nom":"Société Générale","secteur":"Banque",'
        '"cours":14500,"var_1j":"+0.5%","reco":"ACHAT","score":82,'
        '"mm":"haussier","boll":"neutre","macd":"haussier","rsi":"neutre","stoch":"haussier"}\n]\n\n'
        f"TEXTE SOURCE :\n{ticker_context}"
    )

    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            msg = client.messages.create(
                model=_MODEL, max_tokens=2048,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            print(f"  [Extractor/Batch] Tentative {attempt}/{_MAX_RETRIES} — "
                  f"{len(raw)} chars reçus (prompt envoyé : {len(prompt):,} chars)")

            cleaned = clean_json_string(raw)
            parsed  = safe_json_load(cleaned)
            if parsed:
                print(f"  [Extractor/Batch] OK : {len(parsed)} société(s) parsée(s)")
                return parsed

            print(f"  [Extractor/Batch] Résultat vide — tentative {attempt}/{_MAX_RETRIES}")
        except Exception as exc:
            print(f"  [Extractor/Batch] Erreur tentative {attempt} : {exc}")

    print(f"  [Extractor/Batch] Échec total — fallback pour : {tickers}")
    return [
        {"ticker": t, "nom": t, "secteur": None, "cours": None, "var_1j": None,
         "reco": "INCONNU", "score": None,
         "mm": None, "boll": None, "macd": None, "rsi": None, "stoch": None}
        for t in tickers
    ]


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION COMPLÉMENTAIRE : COURS 100J + FONDAMENTAUX
# ═══════════════════════════════════════════════════════════════════════════════

_EXTRA_FIELDS = (
    "cours_debut", "cours_fin", "plus_haut_100j", "plus_bas_100j",
    "perf_100j", "date_debut_100j", "date_fin_100j",
    "ca", "ca_date", "resultat_net", "rn_date",
    "marge_nette", "mn_date", "roe", "roe_date",
    "roa", "roa_date", "dividende", "div_date",
    # Métriques de marché (à afficher dans MÉTRIQUES DE MARCHÉ)
    "capitalisation_boursiere", "capi_date",
    "volume_moyen_30j", "volume_30j_date",
    "nb_actions", "nb_actions_date",
    "rendement_dividende",
    # Ratios financiers supplémentaires (pages 42-66 du rapport source)
    "per", "per_date", "per_sectoriel",
    "dette_cp", "dette_cp_date",
    "dette_nette", "ebitda", "dette_nette_ebitda",
    "croissance_ca", "croissance_ca_date",
    "croissance_ca_1an", "croissance_ca_3ans",
    # Indicateurs techniques : valeurs numériques (en plus du signal)
    "mm20_valeur", "mm50_valeur",
    "boll_inf", "boll_sup",
    "macd_valeur", "macd_signal_line",
    "rsi_valeur",
    "stoch_k", "stoch_d",
    # Évolution historique sur 2-3 ans
    "ca_n_1", "ca_n_2",
    "resultat_n_1", "resultat_n_2",
    "roe_n_1", "roe_n_2",
    # Synthèse financière qualitative
    "forces_financieres", "faiblesses_financieres",
    "synthese_financiere", "date_donnees_financieres",
    # ── Données financières structurées (banques / sociétés cotées) ──
    # BILAN ACTIF
    "caisse_banque_centrale", "effets_publics", "creances_interbancaires",
    "creances_clientele", "immob_incorporelles", "immob_corporelles",
    "tresorerie_actif", "total_actif",
    # BILAN PASSIF
    "capital_souscrit", "reserves", "capitaux_propres", "capitaux_permanents",
    "dettes_interbancaires", "dettes_clientele", "dettes_financieres_totales",
    "dettes_totales",
    # COMPTE DE RÉSULTAT (resultat_net déjà inclus plus haut)
    "pnb", "interets_produits", "interets_charges", "commissions_produits",
    "charges_generales", "charges_personnel", "rbe", "resultat_exploitation",
    "provisions",
    # RATIOS DE RENTABILITÉ
    "marge_operationnelle", "coefficient_exploitation", "cout_du_risque",
    # STRUCTURE & LIQUIDITÉ
    "autonomie_financiere", "dependance_financiere", "gearing",
    "solvabilite_generale", "liquidite_immediate", "couverture_interets",
)


def extract_extra(sections: dict, tickers: list) -> list:
    """
    Extrait les bornes de cours sur 100 jours et les indicateurs fondamentaux
    (CA, résultat net, marge nette, ROE, ROA, dividende) avec leur date.
    Pass séparé du batch principal pour isoler la complexité de cette extraction.
    Retourne list[dict] indexé par ticker.
    """
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    ticker_context = build_ticker_context(sections, tickers)
    tickers_str = ", ".join(tickers)

    prompt = (
        "Tu analyses un rapport boursier BRVM. "
        f"Pour chaque société : {tickers_str}, extrais les données suivantes "
        "(y compris les sections 'Analyse technique', 'Analyse financière', "
        "'Résultats financiers', 'Métriques de marché' si présentes).\n\n"
        "RÈGLES STRICTES :\n"
        "1. Réponds UNIQUEMENT en JSON valide, commençant par [ et terminant par ].\n"
        "2. null pour toute valeur absente — JAMAIS d'invention.\n"
        "3. Cours : nombres en FCFA sans unité ni séparateur (ex: 1440, 14500).\n"
        "4. perf_100j : variation signée avec %, ex: '-7.29%', '+26.95%'.\n"
        "5. Dates au format JJ/MM/AAAA ou AAAA-MM-JJ tel que dans le texte.\n"
        "6. CA, résultat_net, dividende, capitalisation_boursiere, dette_nette, "
        "ebitda : conserve la valeur ET l'unité du texte "
        "(ex: '42,45 milliards FCFA', '22,3 millions FCFA', '150 FCFA').\n"
        "7. marge_nette / roe / roa / croissance_ca / croissance_ca_1an / "
        "croissance_ca_3ans / rendement_dividende : pourcentage avec %, "
        "ex: '6,9%', '1,46%', '+12,3%'.\n"
        "8. per / per_sectoriel : ratio cours/bénéfice (nombre simple, ex: 8.5).\n"
        "9. dette_cp / dette_nette_ebitda : ratio dette/capitaux propres ou "
        "dette nette / EBITDA en valeur ou multiple (ex: '0.45', '1,2x', '2.3x').\n"
        "10. volume_moyen_30j : nombre d'actions échangées en moyenne sur 30 jours "
        "(nombre simple sans unité, ex: 12450).\n"
        "11. nb_actions : nombre total d'actions en circulation (nombre simple, ex: 4500000).\n"
        "12. ca_n_1 / resultat_n_1 / roe_n_1 = exercice précédent (N-1) ; "
        "ca_n_2 / resultat_n_2 / roe_n_2 = exercice N-2 si disponible.\n"
        "13. Indicateurs techniques NUMÉRIQUES :\n"
        "    - mm20_valeur / mm50_valeur : valeurs des moyennes mobiles 20j et 50j en FCFA.\n"
        "    - boll_inf / boll_sup : bornes inférieure et supérieure des bandes de Bollinger.\n"
        "    - macd_valeur / macd_signal_line : valeur du MACD et de sa ligne de signal.\n"
        "    - rsi_valeur : valeur du RSI (nombre entre 0 et 100).\n"
        "    - stoch_k / stoch_d : valeurs %K et %D du stochastique.\n"
        "14. forces_financieres / faiblesses_financieres : tableau de 3 points "
        "max chacun (chaînes courtes <120 chars).\n"
        "15. synthese_financiere : 2-3 phrases résumant la santé financière.\n"
        "16. date_donnees_financieres : date de référence des données financières "
        "(ex: '31/12/2025', 'T3 2025').\n"
        "17. DONNÉES FINANCIÈRES STRUCTURÉES (sections 'Bilan', 'Compte de "
        "résultat', 'Ratios' du rapport — typique des banques BOAC/SGBCI/ETI...) :\n"
        "    BILAN ACTIF (montants avec unité : ex '125,4 Mds FCFA') :\n"
        "      caisse_banque_centrale, effets_publics, creances_interbancaires,\n"
        "      creances_clientele, immob_incorporelles, immob_corporelles,\n"
        "      tresorerie_actif, total_actif.\n"
        "    BILAN PASSIF (montants avec unité) :\n"
        "      capital_souscrit, reserves, capitaux_propres, capitaux_permanents,\n"
        "      dettes_interbancaires, dettes_clientele, dettes_financieres_totales,\n"
        "      dettes_totales.\n"
        "    COMPTE DE RÉSULTAT (montants avec unité) :\n"
        "      pnb (Produit Net Bancaire), interets_produits, interets_charges,\n"
        "      commissions_produits, charges_generales, charges_personnel,\n"
        "      rbe (Résultat Brut d'Exploitation), resultat_exploitation, provisions.\n"
        "    RATIOS DE RENTABILITÉ (pourcentages avec %) :\n"
        "      marge_operationnelle, coefficient_exploitation, cout_du_risque.\n"
        "    STRUCTURE & LIQUIDITÉ (pourcentages avec % ou ratios numériques) :\n"
        "      autonomie_financiere, dependance_financiere, gearing,\n"
        "      solvabilite_generale, liquidite_immediate, couverture_interets.\n"
        "    Toutes ces données : null si absentes — JAMAIS d'invention.\n\n"
        "Schéma exact :\n"
        "[\n"
        '  {\n'
        '    "ticker": "BNBC",\n'
        '    "cours_debut": 1440, "cours_fin": 1335,\n'
        '    "plus_haut_100j": 1785, "plus_bas_100j": 1335,\n'
        '    "perf_100j": "-7.29%",\n'
        '    "date_debut_100j": "2026-02-23", "date_fin_100j": "2026-05-07",\n'
        '    "ca": "42,45 Mds FCFA", "ca_date": "31/12/2025",\n'
        '    "resultat_net": "22,3 M FCFA", "rn_date": "31/12/2025",\n'
        '    "marge_nette": "0,05%", "mn_date": "31/12/2025",\n'
        '    "roe": "5,8%", "roe_date": "31/12/2025",\n'
        '    "roa": "1,2%", "roa_date": "31/12/2025",\n'
        '    "dividende": "150 FCFA", "div_date": "31/12/2025",\n'
        '    "capitalisation_boursiere": "32,1 Mds FCFA", "capi_date": "07/05/2026",\n'
        '    "volume_moyen_30j": 12450, "volume_30j_date": "07/05/2026",\n'
        '    "nb_actions": 24000000, "nb_actions_date": "31/12/2025",\n'
        '    "rendement_dividende": "3,5%",\n'
        '    "per": 8.5, "per_date": "31/12/2025", "per_sectoriel": 11.2,\n'
        '    "dette_cp": "0.45", "dette_cp_date": "31/12/2025",\n'
        '    "dette_nette": "5,2 Mds FCFA", "ebitda": "8,1 Mds FCFA", "dette_nette_ebitda": "0.64x",\n'
        '    "croissance_ca": "+5.2%", "croissance_ca_date": "31/12/2025",\n'
        '    "croissance_ca_1an": "+5,2%", "croissance_ca_3ans": "+14,8%",\n'
        '    "mm20_valeur": 1378, "mm50_valeur": 1412,\n'
        '    "boll_inf": 1305, "boll_sup": 1455,\n'
        '    "macd_valeur": -8.4, "macd_signal_line": -5.1,\n'
        '    "rsi_valeur": 42.5,\n'
        '    "stoch_k": 28.3, "stoch_d": 31.7,\n'
        '    "ca_n_1": "40,3 Mds FCFA", "ca_n_2": "38,1 Mds FCFA",\n'
        '    "resultat_n_1": "20,1 M FCFA", "resultat_n_2": "18,5 M FCFA",\n'
        '    "roe_n_1": "5,8%", "roe_n_2": "5,2%",\n'
        '    "forces_financieres": ["Marge nette stable", "Faible endettement", "Croissance régulière du CA"],\n'
        '    "faiblesses_financieres": ["ROE en baisse", "Dividende non distribué", "Liquidité limitée"],\n'
        '    "synthese_financiere": "Société aux fondamentaux solides malgré une rentabilité en repli. La structure financière reste saine avec un endettement maîtrisé.",\n'
        '    "date_donnees_financieres": "31/12/2025",\n'
        '    "caisse_banque_centrale": "85,2 Mds FCFA", "effets_publics": "142,7 Mds FCFA",\n'
        '    "creances_interbancaires": "67,3 Mds FCFA", "creances_clientele": "920,5 Mds FCFA",\n'
        '    "immob_incorporelles": "2,1 Mds FCFA", "immob_corporelles": "28,4 Mds FCFA",\n'
        '    "tresorerie_actif": "152,8 Mds FCFA", "total_actif": "1 425,6 Mds FCFA",\n'
        '    "capital_souscrit": "30,0 Mds FCFA", "reserves": "78,5 Mds FCFA",\n'
        '    "capitaux_propres": "118,2 Mds FCFA", "capitaux_permanents": "245,7 Mds FCFA",\n'
        '    "dettes_interbancaires": "95,3 Mds FCFA", "dettes_clientele": "1 048,7 Mds FCFA",\n'
        '    "dettes_financieres_totales": "127,5 Mds FCFA", "dettes_totales": "1 307,4 Mds FCFA",\n'
        '    "pnb": "82,5 Mds FCFA", "interets_produits": "65,1 Mds FCFA",\n'
        '    "interets_charges": "21,4 Mds FCFA", "commissions_produits": "18,8 Mds FCFA",\n'
        '    "charges_generales": "18,2 Mds FCFA", "charges_personnel": "22,6 Mds FCFA",\n'
        '    "rbe": "41,7 Mds FCFA", "resultat_exploitation": "35,8 Mds FCFA",\n'
        '    "provisions": "5,9 Mds FCFA",\n'
        '    "marge_operationnelle": "43,4%", "coefficient_exploitation": "49,5%",\n'
        '    "cout_du_risque": "0,6%",\n'
        '    "autonomie_financiere": "8,3%", "dependance_financiere": "91,7%",\n'
        '    "gearing": "1,08", "solvabilite_generale": "1,09",\n'
        '    "liquidite_immediate": "0,75", "couverture_interets": "3,8"\n'
        "  }\n"
        "]\n\n"
        f"TEXTE SOURCE :\n{ticker_context}"
    )

    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            msg = client.messages.create(
                model=_MODEL, max_tokens=8192,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            print(f"  [Extractor/Extra] Tentative {attempt}/{_MAX_RETRIES} — "
                  f"{len(raw)} chars reçus")
            cleaned = clean_json_string(raw)
            parsed = safe_json_load(cleaned)
            if parsed:
                print(f"  [Extractor/Extra] OK : {len(parsed)} société(s) parsée(s)")
                return parsed
            print(f"  [Extractor/Extra] Résultat vide — tentative {attempt}/{_MAX_RETRIES}")
        except Exception as exc:
            print(f"  [Extractor/Extra] Erreur tentative {attempt} : {exc}")

    print(f"  [Extractor/Extra] Échec total — fallback vide pour : {tickers}")
    return [{"ticker": t, **{f: None for f in _EXTRA_FIELDS}} for t in tickers]


def _merge_extra(companies: list, extras: list) -> list:
    """Fusionne par ticker les données extra dans les sociétés existantes."""
    by_ticker = {str(e.get("ticker") or "").strip().upper(): e for e in extras if e.get("ticker")}
    for c in companies:
        t = str(c.get("ticker") or "").strip().upper()
        ex = by_ticker.get(t) or {}
        for f in _EXTRA_FIELDS:
            if c.get(f) is None and ex.get(f) is not None:
                c[f] = ex.get(f)
    return companies


# ═══════════════════════════════════════════════════════════════════════════════
# VALIDATION QUALITÉ
# ═══════════════════════════════════════════════════════════════════════════════

# mm/boll/macd/rsi/stoch/var_1j exclus de la validation car souvent absents
_ESSENTIAL_FIELDS = ("ticker", "nom", "secteur", "cours", "reco", "score")
_CORE_FIELDS = ("ticker", "nom", "secteur", "cours", "var_1j", "reco", "score",
                "mm", "boll", "macd", "rsi", "stoch")


def _log_extraction_stats(companies: list) -> dict:
    """Calcule et affiche les statistiques qualité."""
    if not companies:
        return {"null_ratio": 1.0, "score_nonzero_ratio": 0.0}

    total_all = len(companies) * len(_CORE_FIELDS)
    nulls_all = sum(1 for c in companies for f in _CORE_FIELDS if c.get(f) is None)
    total_ess = len(companies) * len(_ESSENTIAL_FIELDS)
    nulls_ess = sum(1 for c in companies for f in _ESSENTIAL_FIELDS if c.get(f) is None)
    nonzero   = sum(
        1 for c in companies
        if c.get("score") is not None and float(c.get("score") or 0) > 0
    )
    null_ratio  = nulls_ess / total_ess
    score_ratio = nonzero / len(companies)

    print(f"  [Extractor/Stats] Sociétés     : {len(companies)}")
    print(f"  [Extractor/Stats] Champs null (tous)       : {nulls_all}/{total_all} ({nulls_all/total_all*100:.1f}%)")
    print(f"  [Extractor/Stats] Champs null (essentiels) : {nulls_ess}/{total_ess} ({null_ratio*100:.1f}%)")
    print(f"  [Extractor/Stats] Scores > 0   : {nonzero}/{len(companies)} ({score_ratio*100:.1f}%)")
    if companies:
        print(f"  [Extractor/Debug] 1ère société : "
              f"{json.dumps(companies[0], ensure_ascii=False)}")

    return {"null_ratio": null_ratio, "score_nonzero_ratio": score_ratio}


def validate_extraction(companies: list) -> bool:
    """
    Valide la qualité de l'extraction.
    Bloque le pipeline si les données sont insuffisantes.
    """
    if not companies:
        print("  [Extractor/Validation] ÉCHEC : aucune société extraite.")
        return False

    stats    = _log_extraction_stats(companies)
    all_zero = all(
        c.get("score") is None or float(c.get("score") or 0) == 0
        for c in companies
    )

    if all_zero:
        print("  [Extractor/Validation] ÉCHEC : tous les scores sont null/0.")
        return False

    if stats["null_ratio"] > _MAX_NULL_RATIO:
        print(f"  [Extractor/Validation] ÉCHEC : {stats['null_ratio']*100:.1f}% de champs null "
              f"(seuil {_MAX_NULL_RATIO*100:.0f}%).")
        return False

    print(f"  [Extractor/Validation] OK — {stats['score_nonzero_ratio']*100:.1f}% scores non nuls.")
    return True


# ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE COMPLET
# ═══════════════════════════════════════════════════════════════════════════════

def extract_all(full_text: str, freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Pipeline complet d'extraction :
    0. Nettoyage du texte brut
    1. Segmentation en sections (brvm_global / societes / actualites)
    2. Extraction données BRVM globales (section brvm uniquement)
    3. Identification des tickers (section sociétés, max 25k chars)
    4. Extraction par batch avec contexte ciblé par ticker
    5. Validation qualité -> stop si extraction invalide
    """
    print(f"  [Extractor] Texte source brut : {len(full_text):,} chars")

    # ── Étape 0 : Nettoyage
    print("  [Extractor] Étape 0 : Nettoyage...")
    clean = clean_text(full_text)

    # ── Étape 1 : Segmentation
    print("  [Extractor] Étape 1 : Segmentation...")
    sections = split_text_sections(clean)

    # ── Étape 2 : Données BRVM globales
    print("  [Extractor] Étape 2 : Données BRVM globales...")
    brvm_global = extract_brvm_global(sections)
    if brvm_global:
        print(f"  [Extractor] Composite      : {brvm_global.get('brvm_composite', 'N/A')}")
        print(f"  [Extractor] Top opportunité: {brvm_global.get('top_opportunite', 'N/A')}")
        print(f"  [Extractor] Secteur leader : {brvm_global.get('secteur_leader', 'N/A')}")
    else:
        print("  [Extractor] AVERTISSEMENT : données BRVM globales non extraites.")

    # ── Étape 3 : Tickers
    print("  [Extractor] Étape 3 : Identification des tickers...")
    tickers = get_tickers(sections)
    if not tickers:
        print("  [Extractor] ERREUR : aucun ticker identifié — abandon.")
        return []
    print(f"  [Extractor] {len(tickers)} ticker(s) identifié(s) : {tickers[:15]}")

    # ── Étape 4 : Extraction par batch
    all_companies  = []
    total_batches  = (len(tickers) + _BATCH_SIZE - 1) // _BATCH_SIZE
    for i in range(0, len(tickers), _BATCH_SIZE):
        batch     = tickers[i: i + _BATCH_SIZE]
        batch_num = i // _BATCH_SIZE + 1
        print(f"  [Extractor] Étape 4 — Batch {batch_num}/{total_batches} : {', '.join(batch)}")
        companies = extract_batch(sections, batch, freq, period_info)
        print(f"  [Extractor] Batch {batch_num} -> {len(companies)} société(s)")

        print(f"  [Extractor] Étape 4bis — Extra (cours 100j + fondamentaux) batch {batch_num}/{total_batches}")
        extras = extract_extra(sections, batch)
        companies = _merge_extra(companies, extras)

        all_companies.extend(companies)

    print(f"  [Extractor] Total brut : {len(all_companies)} société(s).")

    # ── Étape 5 : Validation
    print("  [Extractor] Étape 5 : Validation qualité...")
    if not validate_extraction(all_companies):
        print("  [Extractor] ERREUR : Extraction failed — pipeline interrompu.")
        return []

    if brvm_global:
        for c in all_companies:
            c["_brvm_global"] = brvm_global

    return all_companies
