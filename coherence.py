"""
coherence.py — Arbitre de cohérence des fiches sociétés BRVM.

Rôle : point UNIQUE de décision qui, avant la génération Word, réconcilie les
valeurs issues des règles Python internes (dict ``s``) avec celles du rapport
source (PARTIE 4 notamment). Objectif : éliminer les contradictions entre
l'en-tête, les métriques de marché, l'analyse technique, le profil de risque et
la conclusion d'investissement d'une même fiche.

Point d'entrée : ``arbitrate(s, parties)`` — mute ``s`` en place et le renvoie.
Aucun appel réseau / LLM : uniquement de l'extraction regex sur des textes déjà
présents en mémoire (les PARTIES 0-4 extraites par ``_extract_parties``).
"""

import re

# Rang de prudence du risque (croissant = plus prudent).
ORDRE_RISQUE = {
    "faible": 0,
    "modéré": 1, "modere": 1, "modérée": 1, "moderee": 1, "moyen": 1, "moyenne": 1,
    "élevé": 2, "eleve": 2, "élevée": 2, "elevee": 2, "fort": 2, "forte": 2,
}

# Rang de confiance (croissant = plus confiant).
ORDRE_CONFIANCE = {
    "faible": 0, "faibl": 0, "basse": 0,
    "modérée": 1, "moderee": 1, "modéré": 1, "modere": 1, "moyen": 1, "moyenne": 1,
    "élevée": 2, "elevee": 2, "élevé": 2, "eleve": 2, "haute": 2, "forte": 2, "fort": 2,
}

# Correspondance vocabulaire source (PARTIE 4) -> clés internes de l'action_map
# utilisée par build_conclusion(). Sans cette table, une reco source telle que
# « CONSERVER » ne matcherait aucune clé et retomberait sur le défaut.
RECO_SOURCE_TO_INTERNE = {
    "ACHAT FORT": "ACHAT FORT", "ACHETER FORT": "ACHAT FORT",
    "RENFORCER": "ACHAT", "ACCUMULER": "ACHAT", "ACHETER": "ACHAT", "ACHAT": "ACHAT",
    "CONSERVER": "SURVEILLER", "MAINTENIR": "SURVEILLER", "GARDER": "SURVEILLER",
    "SURVEILLER": "SURVEILLER", "NEUTRE": "SURVEILLER", "ATTENDRE": "SURVEILLER",
    "ALLÉGER": "PRUDENCE", "ALLEGER": "PRUDENCE", "RÉDUIRE": "PRUDENCE",
    "REDUIRE": "PRUDENCE", "PRUDENCE": "PRUDENCE",
    "VENDRE": "ÉVITER", "SORTIR": "ÉVITER", "ÉVITER": "ÉVITER", "EVITER": "ÉVITER",
}


def _norm(txt):
    return re.sub(r"\s+", " ", str(txt or "")).strip()


def _rang_risque(niveau):
    return ORDRE_RISQUE.get(_norm(niveau).lower(), None)


def _rang_confiance(niveau):
    return ORDRE_CONFIANCE.get(_norm(niveau).lower(), None)


def _to_float(v):
    """Parse robuste FR/EN vers float, ou None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v)).strip()
    if not s or s in ("-", ".", ","):
        return None
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    elif "," in s and "." in s:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# ── Règle 1 — extraction de la décision source (PARTIE 4) ────────────────────

def _extraire_p4(parties):
    """Extrait reco / confiance / risque / horizon depuis le texte PARTIE 4."""
    p4 = _norm((parties or {}).get("p4", ""))
    out = {"reco": None, "confiance": None, "risque": None, "horizon": None}
    if not p4:
        return out

    # Reco : mot(s) après « Recommandation ». Le 2e mot n'est autorisé que s'il
    # s'agit de « FORT » (seule reco composée du référentiel), pour ne pas
    # capturer la suite de phrase (« CONSERVER, confiance ... »).
    m = re.search(r"[Rr]ecommandation\s*(?:source)?\s*[:\-–]?\s*"
                  r"([A-Za-zÀ-ÿ]+)(\s+FORT)?", p4, re.I)
    if m:
        out["reco"] = (m.group(1) + (m.group(2) or "")).strip().upper()

    m = re.search(r"confiance\s*[:\-–]?\s*([A-Za-zÀ-ÿ]+)", p4, re.I)
    if m:
        out["confiance"] = m.group(1).strip()

    # Accepte « risque global Élevé », « risque Élevé », « risque : élevé ».
    m = re.search(r"risque(?:\s+global)?\s*[:\-–]?\s*"
                  r"(faibles?|mod[ée]r[ée]e?|moyenne?|[ée]lev[ée]e?)", p4, re.I)
    if m:
        out["risque"] = m.group(1).strip()

    m = re.search(r"horizon[^.\n]*?"
                  r"(\d+\s*[-–à]\s*\d+\s*mois|>\s*\d+\s*mois|<\s*\d+\s*mois|\d+\s*mois)",
                  p4, re.I)
    if m:
        out["horizon"] = _norm(m.group(1))
    return out


def _purge_defensif(s):
    """Retire les étiquettes « profil défensif solide » devenues trompeuses des
    listes de facteurs de risque (là où elles sont réellement affichées)."""
    motif = re.compile(r"profil\s+d[ée]fensif\s+solide|d[ée]fensif\s+solide|"
                       r"profil\s+(?:tr[èe]s\s+)?d[ée]fensif", re.I)
    for key in ("risques", "faiblesses_financieres"):
        val = s.get(key)
        if isinstance(val, list):
            s[key] = [x for x in val if not motif.search(str(x))]


def arbitrate(s, parties):
    """Réconcilie ``s`` (règles internes) avec la PARTIE 4 source. Mute ``s``."""
    src = _extraire_p4(parties)

    # ── Règle 1 — la PARTIE 4 pilote la conclusion (si elle existe) ──────────
    decision_systeme = (_norm(s.get("decision")) or _norm(s.get("reco"))
                        or "SURVEILLER").upper()
    if src["reco"]:
        reco_source = src["reco"]
        interne = RECO_SOURCE_TO_INTERNE.get(reco_source, "SURVEILLER")
        s["decision_systeme"] = decision_systeme
        s["decision_source"] = reco_source
        s["decision_final"] = interne                 # compatible action_map
        s["reco_display"] = f"{reco_source} (source)"
        s["reco_transparence"] = (
            f"Reco système (règles) : {decision_systeme} — "
            f"Reco retenue (rapport source) : {reco_source}"
        )
        if src["horizon"]:
            s["horizon_final"] = src["horizon"]
        # Confiance : retenir la plus PRUDENTE (rang le plus bas).
        if src["confiance"]:
            r_src = _rang_confiance(src["confiance"])
            r_cur = _rang_confiance(s.get("confiance"))
            if r_src is not None and (r_cur is None or r_src < r_cur):
                s["confiance"] = src["confiance"].capitalize()

    # ── Règle 2 — le risque retenu = le PLUS prudent des deux ───────────────
    r_regles = _rang_risque(s.get("risque"))
    r_source = _rang_risque(src["risque"])
    if r_source is not None and (r_regles is None or r_source > r_regles):
        s["risque"] = src["risque"].lower()
        s["risque_arbitre"] = True

    # ── Règle 3 — détection de divergence technique ─────────────────────────
    rsi = _to_float(s.get("rsi_valeur"))
    stoch_k = _to_float(s.get("stoch_k"))
    stoch_d = _to_float(s.get("stoch_d"))
    surachat = ((rsi is not None and rsi > 70)
                or (stoch_k is not None and stoch_k > 80)
                or (stoch_d is not None and stoch_d > 80))
    survente = ((rsi is not None and rsi < 30)
                or (stoch_k is not None and stoch_k < 20)
                or (stoch_d is not None and stoch_d < 20))

    def _has(v, mot):
        return mot in str(v or "").lower()

    nb_haussier = sum(1 for k in ("mm", "macd", "boll") if _has(s.get(k), "haussier"))
    nb_baissier = sum(1 for k in ("mm", "macd", "boll") if _has(s.get(k), "baissier"))
    tendance_haussiere = nb_haussier >= 2
    tendance_baissiere = nb_baissier >= 2

    div_actuelle = _norm(s.get("divergence")).lower()
    a_deja_div = div_actuelle not in ("", "aucune", "—", "none", "non")

    if surachat and tendance_haussiere and not a_deja_div:
        s["divergence"] = ("⚠ RSI/Stochastique en zone de surachat malgré une "
                           "tendance haussière MM/MACD → risque de consolidation "
                           "avant poursuite")
        s["signal_label_override"] = "DOMINANTE HAUSSIÈRE — SIGNAUX DE SURACHAT"
        s["signal_override_kind"] = "surachat"
    elif survente and tendance_baissiere and not a_deja_div:
        s["divergence"] = ("⚠ RSI/Stochastique en zone de survente malgré une "
                           "tendance baissière MM/MACD → possible rebond technique")
        s["signal_label_override"] = "DOMINANTE BAISSIÈRE — SIGNAUX DE SURVENTE"
        s["signal_override_kind"] = "survente"

    # ── Règle 5 — neutraliser un « profil défensif » contradictoire ─────────
    r_final = _rang_risque(s.get("risque"))
    if r_final is not None and r_final >= 1:
        _purge_defensif(s)

    return s
