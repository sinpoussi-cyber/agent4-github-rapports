"""
Couche d'enrichissement Python.
Responsabilité unique : dériver tous les champs métier et textes narratifs
depuis les données minimales (ticker, nom, secteur, cours, var_1j, reco, score).
Aucun appel LLM.
"""


# ── Règles métier ─────────────────────────────────────────────────────────────

def _score_float(score) -> float:
    try:
        return float(score or 0)
    except (ValueError, TypeError):
        return 0.0


def _decision(score: float, reco: str) -> str:
    r = str(reco or "").upper()
    achat = any(w in r for w in ("ACHAT", "BUY"))
    vente = any(w in r for w in ("VENTE", "SELL"))
    if score >= 75 and achat:
        return "ACHAT FORT"
    if score >= 60 and achat:
        return "ACHAT"
    if score >= 60 and not vente:
        return "SURVEILLER"
    if score >= 40:
        return "PRUDENCE"
    return "ÉVITER"


def _risque(score: float) -> str:
    if score >= 70:
        return "faible"
    if score >= 50:
        return "modéré"
    return "élevé"


def _score_label(score: float) -> str:
    if score >= 75:
        return "Excellent"
    if score >= 60:
        return "Bon"
    if score >= 40:
        return "Moyen"
    return "Faible"


def _liquidite(score: float) -> str:
    if score >= 65:
        return "haute"
    if score >= 40:
        return "moyenne"
    return "faible"


def _stabilite(score: float) -> str:
    if score >= 65:
        return "bonne"
    if score >= 45:
        return "modérée"
    return "fragile"


def _confiance(score: float) -> str:
    if score >= 70:
        return "Élevée"
    if score >= 50:
        return "Modérée"
    return "Faible"


def _reco_src(reco: str) -> str:
    r = str(reco or "").upper()
    if any(w in r for w in ("ACHAT", "BUY")):
        return "vert"
    if any(w in r for w in ("VENTE", "SELL")):
        return "rouge"
    return "orange"


def _volatilite(score: float) -> str:
    if score >= 70:
        return "faible"
    if score >= 50:
        return "modérée"
    return "élevée"


def _signal_global(score: float) -> str:
    if score >= 65:
        return "haussier"
    if score >= 45:
        return "neutre"
    return "baissier"


def _tendance(score: float) -> str:
    if score >= 65:
        return "haussière"
    if score >= 45:
        return "neutre"
    return "baissière"


def _risques_list(score: float, reco: str) -> list:
    r = str(reco or "").upper()
    if score >= 70:
        return ["Risque limité", "Profil défensif solide"]
    if score >= 50:
        return ["Volatilité modérée", "Dépendance aux conditions de marché"]
    base = ["Risque élevé", "Momentum défavorable"]
    if "VENTE" in r or "SELL" in r:
        base.append("Pression vendeuse persistante")
    return base


def _perspectives(score: float, reco: str, decision: str) -> str:
    r = str(reco or "").upper()
    if score >= 70 and "ACHAT" in r:
        return "Positives à moyen terme"
    if score >= 60:
        return "Constructives, à surveiller"
    if score >= 40:
        return "Neutres à court terme — réévaluation recommandée"
    return "Prudence à court terme — risque de détérioration"


# ── Génération des textes narratifs ──────────────────────────────────────────

def _analyse_cours_100j(ticker: str, score: float, reco: str, decision: str, signal: str) -> str:
    score_label = _score_label(score)
    return (
        f"{ticker} affiche un score global de {score:.0f}/100 ({score_label}), "
        f"avec une recommandation {reco} et une décision d'investissement : {decision}. "
        f"L'orientation générale des indicateurs est {signal}, "
        f"cohérente avec le profil de risque {_risque(score)} identifié sur ce titre. "
        f"Les données de cours détaillées seront intégrées lors de la prochaine mise à jour complète."
    )


def _indicator_signal(score: float) -> str:
    return _signal_global(score)


def _indicator_detail(label: str, signal: str) -> str:
    templates = {
        "haussier":  f"L'indicateur {label} est positionné favorablement, soutenant la tendance haussière.",
        "neutre":    f"L'indicateur {label} évolue en zone neutre, sans signal directionnel fort.",
        "baissier":  f"L'indicateur {label} indique une pression baissière — surveiller l'évolution.",
    }
    return templates.get(signal, f"Signal {label} : {signal}.")


def _synthese_tech(score: float, decision: str, signal: str) -> str:
    return (
        f"Score technique global : {score:.0f}/100 — Décision : {decision}. "
        f"Les indicateurs convergent vers un signal {signal}. "
        f"Profil de risque {_risque(score)} — recommandation maintenue."
    )


def _analyse_fond(ticker: str, score: float, reco: str, score_label: str, decision: str) -> str:
    return (
        f"{ticker} présente un profil fondamental {score_label} (score {score:.0f}/100). "
        f"La recommandation analytique est {reco} — position : {decision}. "
        f"Liquidité {_liquidite(score)}, stabilité {_stabilite(score)}. "
        f"Une analyse approfondie des publications récentes sera intégrée lors du prochain rapport détaillé."
    )


# ── Point d'entrée ────────────────────────────────────────────────────────────

def enrich(companies: list) -> list:
    """
    Enrichit chaque société avec tous les champs métier dérivés de ticker/reco/score.
    Retourne la liste enrichie, prête pour la génération Word.
    """
    enriched = []
    for c in companies:
        ticker = str(c.get("ticker") or "").strip()
        if not ticker:
            print(f"  [Enricher] SKIP société sans ticker : {c}")
            continue

        reco = str(c.get("reco") or "NEUTRE").strip()
        score = _score_float(c.get("score"))
        signal = _signal_global(score)
        decision = _decision(score, reco)
        score_label = _score_label(score)
        nom = str(c.get("nom") or ticker).strip()

        e = {
            # Champs LLM préservés
            **c,
            # Champs métier dérivés
            "reco":       reco,
            "score":      score,
            "decision":   decision,
            "score_label": score_label,
            "risque":     c.get("risque") or _risque(score),
            "liquidite":  c.get("liquidite") or _liquidite(score),
            "stabilite":  c.get("stabilite") or _stabilite(score),
            "confiance":  c.get("confiance") or _confiance(score),
            "divergence": c.get("divergence") or "aucune",
            "reco_src":   c.get("reco_src") or _reco_src(reco),
            "volatilite": c.get("volatilite") or _volatilite(score),
            "beta":       c.get("beta") or "1.0",
            "tendance_100j": c.get("tendance_100j") or _tendance(score),
            # Signaux indicateurs techniques
            "mm":    c.get("mm") or signal,
            "boll":  c.get("boll") or signal,
            "macd":  c.get("macd") or signal,
            "rsi":   c.get("rsi") or ("neutre" if score < 70 else "élevé"),
            "stoch": c.get("stoch") or signal,
            # Détails indicateurs
            "mm_signal":    c.get("mm_signal") or signal.capitalize(),
            "boll_signal":  c.get("boll_signal") or signal.capitalize(),
            "macd_signal":  c.get("macd_signal") or signal.capitalize(),
            "rsi_signal":   c.get("rsi_signal") or ("Neutre" if score < 70 else "Élevé"),
            "stoch_signal": c.get("stoch_signal") or signal.capitalize(),
            "mm_detail":    c.get("mm_detail") or _indicator_detail("Moyennes Mobiles", signal),
            "boll_detail":  c.get("boll_detail") or _indicator_detail("Bandes de Bollinger", signal),
            "macd_detail":  c.get("macd_detail") or _indicator_detail("MACD", signal),
            "rsi_detail":   c.get("rsi_detail") or _indicator_detail("RSI", signal),
            "stoch_detail": c.get("stoch_detail") or _indicator_detail("Stochastique", signal),
            # Textes narratifs
            "analyse_cours_100j": (
                c.get("analyse_cours_100j")
                or _analyse_cours_100j(ticker, score, reco, decision, signal)
            ),
            "synthese_tech": (
                c.get("synthese_tech")
                or _synthese_tech(score, decision, signal)
            ),
            "analyse_fond": (
                c.get("analyse_fond")
                or _analyse_fond(nom, score, reco, score_label, decision)
            ),
            "analyse_fond_recente": c.get("analyse_fond_recente"),
            # Listes et perspectives
            "risques":      c.get("risques") or _risques_list(score, reco),
            "perspectives": c.get("perspectives") or _perspectives(score, reco, decision),
            "resume_rapport": c.get("resume_rapport") or f"{score_label} — {decision}",
        }
        enriched.append(e)

    print(f"  [Enricher] {len(enriched)} société(s) enrichie(s) sur {len(companies)} en entrée.")
    return enriched
