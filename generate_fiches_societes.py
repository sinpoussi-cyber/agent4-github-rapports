import io
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from extractor import extract_all
from enricher import enrich

_MARGIN_CM = 1.5


# ── Helpers bas niveau ────────────────────────────────────────────────────────

def _rgb(hex6: str) -> RGBColor:
    return RGBColor(int(hex6[:2], 16), int(hex6[2:4], 16), int(hex6[4:], 16))


def _cell_bg(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _cell_margins(cell, twips: int = 40):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side in ("top", "bottom", "left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(twips))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _cw(cell, text, bold=False, size=8, color=None, bg=None, align=None):
    if bg:
        _cell_bg(cell, bg)
    _cell_margins(cell)
    para = cell.paragraphs[0]
    para.clear()
    run = para.add_run(str(text) if text is not None else "")
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = _rgb(color)
    if align is not None:
        para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    return run


def _cp(p, before=0, after=1):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)


def _narrative(doc, text: str, size: int = 9, italic: bool = False, color: str = None):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.italic = italic
    if color:
        r.font.color.rgb = _rgb(color)
    return p


# ── Helpers visuels ───────────────────────────────────────────────────────────

def _section_heading(doc, text: str, color: str = "1A73E8"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = _rgb(color)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "4")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pBdr.append(bot)
    pPr.append(pBdr)


def _sub_heading(doc, text: str, color: str = "444444"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb(color)


def _key_bloc(doc, label: str, texte: str, bg: str, fg: str = "1A1A1A"):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.rows[0].cells[0]
    _cell_bg(cell, bg)
    _cell_margins(cell, 80)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r1 = p.add_run(f"{label}  ")
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = _rgb(fg)
    r2 = p.add_run(texte)
    r2.font.size = Pt(9)
    r2.font.color.rgb = _rgb("333333")
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(3)


def _add_separator(doc, color: str = "CCCCCC"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "4")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pBdr.append(bot)
    pPr.append(pBdr)


# ── Helpers métier ────────────────────────────────────────────────────────────

def _signal_emoji(signal: str) -> str:
    s = str(signal or "").lower()
    if any(w in s for w in ("haussier", "positif", "achat", "fort", "bon", "élevé", "eleve")):
        return "🟢"
    if any(w in s for w in ("baissier", "négatif", "negatif", "vente", "faible")):
        return "🔴"
    return "🟡"


def _signal_bg(signal: str) -> str:
    s = str(signal or "").lower()
    if any(w in s for w in ("haussier", "positif", "achat", "élevé")):
        return "C6EFCE"
    if any(w in s for w in ("baissier", "négatif", "vente")):
        return "FFC7CE"
    return "FFEB9C"


def _score_label_color(score) -> tuple:
    try:
        s = float(score or 0)
    except (ValueError, TypeError):
        s = 0.0
    if s >= 75:
        return "Excellent", "0F9D58"
    if s >= 60:
        return "Bon", "1A73E8"
    if s >= 40:
        return "Moyen", "E37400"
    return "Faible", "D93025"


def _reco_bg(reco: str) -> str:
    r = str(reco).upper()
    if "ACHAT" in r:
        return "C6EFCE"
    if "VENTE" in r:
        return "FFC7CE"
    return "FFEB9C"


def _reco_fg(reco: str) -> str:
    r = str(reco).upper()
    if "ACHAT" in r:
        return "0F9D58"
    if "VENTE" in r:
        return "D93025"
    return "E37400"


def _risque_bg(risque: str) -> str:
    r = str(risque or "").lower()
    if "faible" in r:
        return "C6EFCE"
    if "élevé" in r or "eleve" in r:
        return "FFC7CE"
    return "FFEB9C"


def _var_color(var: str) -> str:
    v = str(var or "").strip()
    if v.startswith("+"):
        return "EBF7EE"
    if v.startswith("-"):
        return "FDEEEE"
    return "FFFFFF"


def _s(data, key, default="") -> str:
    v = data.get(key)
    return str(v) if v is not None else default


def _sl(data, key) -> list:
    v = data.get(key)
    return v if isinstance(v, list) else []


def _score_f(data) -> float:
    try:
        return float(data.get("score") or 0)
    except (ValueError, TypeError):
        return 0.0


# ── Extraction texte source ───────────────────────────────────────────────────

def _extract_text(doc_bytes: bytes) -> str:
    doc = Document(io.BytesIO(doc_bytes))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(" | ".join(cells))
    return "\n".join(lines)


# ── Contexte multi-documents ─────────────────────────────────────────────────

def _build_context(docs_bytes: list, freq: str) -> str:
    if len(docs_bytes) == 1:
        return _extract_text(docs_bytes[0])

    max_older = {"HEBDO": 6, "MENSUEL": 9, "TRIM": 12, "ANNUEL": 14}.get(freq, 5)
    chars_older = {"HEBDO": 3000, "MENSUEL": 2000, "TRIM": 1500, "ANNUEL": 1000}.get(freq, 2000)

    recent = _extract_text(docs_bytes[0])[:25000]
    older_parts = []
    for i, db in enumerate(docs_bytes[1:max_older + 1], 1):
        excerpt = _extract_text(db)[:chars_older]
        older_parts.append(f"--- Document J-{i} ---\n{excerpt}")

    return (
        "=== DOCUMENT LE PLUS RÉCENT (J) ===\n"
        + recent
        + ("\n\n=== HISTORIQUE (extraits) ===\n" + "\n\n".join(older_parts) if older_parts else "")
    )


# ═══════════════════════════════════════════════════════════════════════════════
# FONCTIONS DE CONSTRUCTION DE LA FICHE
# ═══════════════════════════════════════════════════════════════════════════════

def build_header(doc, s: dict, date_str: str):
    """
    En-tête complet de la fiche :
    Ligne 1 — Ticker | Nom | Secteur | Date
    Ligne 2 — Score/Label | Recommandation | Indicateurs (🟢🟡🔴)
    Ligne 3 — Décision | Confiance | Tendance | Résumé
    """
    ticker = _s(s, "ticker", "???")
    nom = _s(s, "nom", ticker)
    secteur = _s(s, "secteur", "—")
    score = s.get("score")
    score_str = f"{score:.0f}" if score is not None else "—"
    score_label, score_color = _score_label_color(score)
    reco = _s(s, "reco", "NEUTRE")
    decision = _s(s, "decision", "—")
    confiance = _s(s, "confiance", "—")
    tendance = _s(s, "tendance_100j", "—")
    resume = _s(s, "resume_rapport", "—")

    tbl = doc.add_table(rows=3, cols=4)
    tbl.style = "Table Grid"

    # ── Ligne 1 : identité
    r0 = tbl.rows[0]
    _cw(r0.cells[0], ticker, bold=True, size=15, color="FFFFFF", bg="1A237E")
    _cw(r0.cells[1], nom[:48], size=9, color="E8EAF6", bg="1A237E")
    _cw(r0.cells[2], secteur, size=8, color="C5CAE9", bg="283593")
    _cw(r0.cells[3], f"Rapport du {date_str}", size=8, color="C5CAE9", bg="283593",
        align=WD_ALIGN_PARAGRAPH.RIGHT)

    # ── Ligne 2 : score + reco + indicateurs
    r1 = tbl.rows[1]
    _cw(r1.cells[0], f"Score : {score_str}/100  —  {score_label}",
        bold=True, size=9, color=score_color, bg="E8EAF6")
    _cw(r1.cells[1], reco, bold=True, size=11,
        color=_reco_fg(reco), bg=_reco_bg(reco))
    ind = "   ".join([
        f"{_signal_emoji(_s(s, 'mm'))} MM",
        f"{_signal_emoji(_s(s, 'boll'))} Boll",
        f"{_signal_emoji(_s(s, 'macd'))} MACD",
        f"{_signal_emoji(_s(s, 'rsi'))} RSI",
        f"{_signal_emoji(_s(s, 'stoch'))} Stoch",
    ])
    r1.cells[2].merge(r1.cells[3])
    _cw(r1.cells[2], ind, size=9, bg="E8EAF6")

    # ── Ligne 3 : décision + confiance + tendance + résumé
    r2 = tbl.rows[2]
    _cw(r2.cells[0], f"Décision : {decision}", bold=True, size=8, bg="F0F4FF")
    _cw(r2.cells[1], f"Confiance : {confiance}", size=8, bg="F5F5F5")
    _cw(r2.cells[2], f"Tendance 100j : {tendance}", size=8, bg="F5F5F5")
    _cw(r2.cells[3], f"Résumé : {resume}", size=8, bg="F5F5F5")


def build_market_table(doc, s: dict):
    """
    Tableau des métriques de marché en 4×2 (8 indicateurs).
    Code couleur sur variation, risque et stabilité.
    """
    _section_heading(doc, "MÉTRIQUES DE MARCHÉ")

    cours = _s(s, "cours") or "—"
    var_1j = _s(s, "var_1j") or "—"
    volatilite = _s(s, "volatilite") or "—"
    beta = _s(s, "beta") or "—"
    liquidite = _s(s, "liquidite") or "—"
    risque = _s(s, "risque") or "—"
    divergence = _s(s, "divergence") or "aucune"
    stabilite = _s(s, "stabilite") or "—"

    risque_bg = _risque_bg(risque)
    var_bg = _var_color(var_1j)
    stab_bg = "C6EFCE" if "bonne" in str(stabilite).lower() else (
        "FFC7CE" if "fragile" in str(stabilite).lower() else "FFEB9C"
    )
    div_bg = "FFEB9C" if divergence.lower() not in ("aucune", "—", "") else "FFFFFF"

    pairs = [
        ("Cours actuel (FCFA)",   cours,      "F0F4FF"),
        ("Variation 1 journée",   var_1j,     var_bg),
        ("Volatilité",            volatilite, "FFFFFF"),
        ("Bêta",                  beta,       "FFFFFF"),
        ("Liquidité",             liquidite,  "FFFFFF"),
        ("Niveau de risque",      risque,     risque_bg),
        ("Divergence tech/fond",  divergence, div_bg),
        ("Stabilité",             stabilite,  stab_bg),
    ]

    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = "Table Grid"
    for i in range(4):
        ll, lv, lbg = pairs[i]
        rl, rv, rbg = pairs[i + 4]
        _cw(tbl.rows[i].cells[0], ll, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[1], lv, size=8, bg=lbg)
        _cw(tbl.rows[i].cells[2], rl, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[3], rv, size=8, bg=rbg)


def build_chart_comment(doc, s: dict):
    """
    Graphique placeholder + commentaire analytique de la courbe.
    Couvre : tendance, volatilité, momentum, phase de marché.
    """
    # ── Placeholder graphique
    p_ph = doc.add_paragraph()
    p_ph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ph.paragraph_format.space_before = Pt(8)
    p_ph.paragraph_format.space_after = Pt(4)
    r_ph = p_ph.add_run(
        "[Courbe d'évolution du cours de l'action sur les 100 derniers jours avec prédictions]"
    )
    r_ph.italic = True
    r_ph.font.size = Pt(9)
    r_ph.font.color.rgb = _rgb("888888")

    # ── Section commentaire
    _section_heading(doc, "COMMENTAIRE DE LA COURBE — ÉVOLUTION 100 JOURS")

    ticker = _s(s, "ticker", "ce titre")
    cours = _s(s, "cours") or "—"
    var_1j = _s(s, "var_1j") or "—"
    tendance = _s(s, "tendance_100j", "neutre")
    volatilite = _s(s, "volatilite", "modérée")

    # Texte principal : utiliser l'analyse enricher si disponible
    analyse = _s(s, "analyse_cours_100j")
    if analyse and analyse not in ("", "—"):
        _narrative(doc, analyse)

    # Interprétation systématique : phase + momentum + implication
    t = str(tendance).lower()
    if "haussier" in t or "hausse" in t:
        phase = "accumulation progressive"
        momentum = "positif et soutenu"
        implication = (
            "Le titre progresse dans une structure haussière avec des corrections limitées. "
            "Les creux successifs sont plus hauts que les précédents — signal de demande structurelle. "
            "Le risque principal est une prise de bénéfices après extension rapide."
        )
    elif "baissier" in t or "baisse" in t:
        phase = "correction / distribution"
        momentum = "négatif et persistant"
        implication = (
            "Le titre est en phase de distribution : les vendeurs dominent les acheteurs. "
            "Chaque rebond technique est suivi d'un nouveau plus bas — structure baissière intacte. "
            "Un signal de retournement confirmé est nécessaire avant tout renforcement."
        )
    else:
        phase = "consolidation latérale"
        momentum = "neutre — sans direction"
        implication = (
            "Le titre évolue en range horizontal. "
            "Cette phase d'équilibre entre acheteurs et vendeurs précède souvent une rupture directionnelle. "
            "Surveiller le franchissement en clôture des bornes du range pour identifier le prochain mouvement."
        )

    _narrative(doc,
               f"Phase de marché : {phase}. Momentum : {momentum}. "
               f"Volatilité sur la période : {volatilite}. "
               f"Cours actuel : {cours} FCFA — variation séance : {var_1j}.")
    _narrative(doc, implication)


def build_technical_analysis(doc, s: dict):
    """
    Analyse technique structurée :
    - Tableau des 5 indicateurs (signal + appréciation + détail)
    - Synthèse globale
    - Évaluation convergence / divergence des signaux
    """
    _section_heading(doc, "ANALYSE TECHNIQUE")

    # ── Tableau des indicateurs
    indicateurs = [
        ("Moyennes Mobiles (MM)", "mm",   "mm_signal",   "mm_detail"),
        ("Bandes de Bollinger",   "boll", "boll_signal", "boll_detail"),
        ("MACD",                  "macd", "macd_signal", "macd_detail"),
        ("RSI",                   "rsi",  "rsi_signal",  "rsi_detail"),
        ("Stochastique",          "stoch","stoch_signal","stoch_detail"),
    ]

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Indicateur", "Appréciation", "Signal", "Analyse détaillée"]):
        _cw(tbl.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")

    for label, sig_key, sig_label_key, detail_key in indicateurs:
        signal = _s(s, sig_key)
        sig_label = _s(s, sig_label_key) or (signal.capitalize() if signal else "—")
        detail = _s(s, detail_key) or "—"
        emoji = _signal_emoji(signal)

        row = tbl.add_row()
        _cw(row.cells[0], f"{emoji}  {label}", bold=True, size=8, bg="EBF0FA")
        _cw(row.cells[1], sig_label, size=8, bg=_signal_bg(signal))
        _cw(row.cells[2], str(signal).upper() if signal else "—", size=8)
        _cw(row.cells[3], detail[:120] if detail != "—" else "—", size=8)

    doc.add_paragraph()

    # ── Synthèse technique (texte enricher)
    synthese = _s(s, "synthese_tech") or _s(s, "analyse_tech")
    if synthese and synthese not in ("", "—"):
        _narrative(doc, synthese, italic=True)

    # ── Évaluation convergence des signaux
    signals_raw = [_s(s, k) for k in ("mm", "boll", "macd", "rsi", "stoch")]
    pos = sum(1 for sg in signals_raw if any(
        w in str(sg).lower() for w in ("haussier", "positif", "achat", "élevé")))
    neg = sum(1 for sg in signals_raw if any(
        w in str(sg).lower() for w in ("baissier", "négatif", "negatif", "vente")))
    neu = len(signals_raw) - pos - neg

    if pos >= 4:
        conv_label = "CONVERGENCE HAUSSIÈRE"
        conv_text = (
            f"{pos}/5 indicateurs en signal positif. "
            "Les signaux techniques sont fortement alignés — biais acheteur confirmé. "
            "Risque principal : surextension si la progression est trop rapide."
        )
        conv_bg, conv_fg = "C6EFCE", "0F9D58"
    elif neg >= 4:
        conv_label = "CONVERGENCE BAISSIÈRE"
        conv_text = (
            f"{neg}/5 indicateurs en signal négatif. "
            "Les signaux techniques convergent vers un biais vendeur. "
            "Éviter toute entrée avant confirmation d'un signal de retournement."
        )
        conv_bg, conv_fg = "FFC7CE", "D93025"
    elif pos > neg:
        conv_label = "DOMINANTE HAUSSIÈRE"
        conv_text = (
            f"{pos} signaux positifs / {neg} négatifs / {neu} neutres. "
            "Signaux partiellement alignés. "
            "Surveiller la confirmation par le volume avant renforcement."
        )
        conv_bg, conv_fg = "E8F8F0", "155724"
    elif neg > pos:
        conv_label = "DOMINANTE BAISSIÈRE"
        conv_text = (
            f"{neg} signaux négatifs / {pos} positifs / {neu} neutres. "
            "Prudence recommandée. "
            "Ne pas renforcer tant que le bilan des signaux ne s'améliore pas."
        )
        conv_bg, conv_fg = "FFF0E6", "C0392B"
    else:
        conv_label = "SIGNAUX MIXTES"
        conv_text = (
            f"{pos} positifs / {neg} négatifs / {neu} neutres. "
            "Absence de convergence claire — pas de biais directionnel dominant. "
            "Attendre un signal confirmé avant toute décision."
        )
        conv_bg, conv_fg = "FFEB9C", "7D6608"

    _key_bloc(doc, f"SYNTHÈSE TECHNIQUE — {conv_label} :", conv_text, conv_bg, conv_fg)


def build_fundamental_analysis(doc, s: dict):
    """
    Analyse fondamentale :
    - Profil de la société (secteur, positionnement, score)
    - Analyse qualitative (texte enricher ou fallback)
    - Profil financier (liquidité, stabilité, confiance)
    - Risques identifiés
    - Perspectives
    """
    _section_heading(doc, "ANALYSE FONDAMENTALE")

    ticker = _s(s, "ticker", "—")
    nom = _s(s, "nom", ticker)
    secteur = _s(s, "secteur", "—")
    score_f = _score_f(s)
    score_str = f"{score_f:.0f}" if s.get("score") is not None else "—"
    score_label, _ = _score_label_color(score_f)
    liquidite = _s(s, "liquidite", "—")
    stabilite = _s(s, "stabilite", "—")
    confiance = _s(s, "confiance", "—")
    reco_src = _s(s, "reco_src", "—")

    # ── Analyse principale
    analyse = _s(s, "analyse_fond_recente")
    if not analyse or analyse in ("", "—", "null", "None"):
        analyse = _s(s, "analyse_fond")

    if analyse and analyse not in ("", "—", "null", "None"):
        _narrative(doc, analyse)
    else:
        _narrative(doc,
                   f"{nom} ({ticker}), cotée dans le secteur {secteur} de la BRVM, "
                   f"présente un profil fondamental {score_label} avec un score composite de {score_str}/100. "
                   f"Liquidité : {liquidite}. Stabilité financière : {stabilite}. "
                   f"Confiance analytique : {confiance}. "
                   "Les données fondamentales détaillées seront intégrées lors du prochain rapport complet.")

    # ── Profil financier synthétique (ligne 1 colonne)
    _sub_heading(doc, "Profil financier")
    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Liquidité", "Stabilité", "Confiance analytique", "Source reco"]):
        _cw(tbl.rows[0].cells[i], h, bold=True, size=8, bg="EBF0FA")
    vals = [liquidite, stabilite, confiance, reco_src]
    bgs = [
        "C6EFCE" if "haute" in str(liquidite).lower() else ("FFC7CE" if "faible" in str(liquidite).lower() else "FFEB9C"),
        _risque_bg(stabilite.replace("bonne", "faible").replace("fragile", "élevé")),
        "C6EFCE" if "élevée" in str(confiance).lower() else ("FFC7CE" if "faible" in str(confiance).lower() else "FFEB9C"),
        "FFFFFF",
    ]
    for i, (v, b) in enumerate(zip(vals, bgs)):
        _cw(tbl.rows[1].cells[i], v, size=8, bg=b)
    doc.add_paragraph()

    # ── Risques
    risques = _sl(s, "risques")
    if risques:
        _key_bloc(doc,
                  "⚠  RISQUES IDENTIFIÉS :",
                  "  |  ".join(str(r) for r in risques[:3]),
                  "FFF0E6", "C0392B")

    # ── Perspectives
    persp = _s(s, "perspectives")
    if persp and persp not in ("", "—"):
        _key_bloc(doc, "PERSPECTIVES :", persp, "E8F8F0", "155724")


def build_conclusion(doc, s: dict):
    """
    Conclusion d'investissement :
    1. Matrice Risque × Horizon de placement
    2. Divergences majeures
    3. Recommandation finale avec action claire
    """
    _section_heading(doc, "CONCLUSION D'INVESTISSEMENT")

    ticker = _s(s, "ticker", "—")
    score_f = _score_f(s)
    score_str = f"{score_f:.0f}" if s.get("score") is not None else "—"
    score_label, score_color = _score_label_color(score_f)
    reco = _s(s, "reco", "NEUTRE")
    decision = _s(s, "decision", "SURVEILLER")
    risque = _s(s, "risque", "modéré")
    divergence = _s(s, "divergence", "aucune")
    confiance = _s(s, "confiance", "Modérée")
    resume = _s(s, "resume_rapport", "—")

    # ── Calcul horizon à partir du score
    if score_f >= 70:
        horizon = "Long terme (> 12 mois)"
        profil_inv = "Investissement core / portefeuille défensif"
    elif score_f >= 55:
        horizon = "Moyen terme (6-12 mois)"
        profil_inv = "Opportunité de portage / portefeuille équilibré"
    elif score_f >= 40:
        horizon = "Court terme (3-6 mois)"
        profil_inv = "Position tactique / portefeuille offensif"
    else:
        horizon = "Très court terme ou abstention"
        profil_inv = "Profil spéculatif — risque élevé"

    # ── 1. Matrice Risque × Horizon
    _sub_heading(doc, "1.  Matrice Risque × Horizon de placement")

    tbl = doc.add_table(rows=2, cols=5)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Valeur", "Risque", "Horizon recommandé", "Profil investisseur", "Score"]):
        _cw(tbl.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")

    risque_bg = _risque_bg(risque)
    _cw(tbl.rows[1].cells[0], ticker, bold=True, size=9, bg="E8EAF6")
    _cw(tbl.rows[1].cells[1], risque.capitalize(), bold=True, size=9, bg=risque_bg)
    _cw(tbl.rows[1].cells[2], horizon, size=9, bg="F5F5F5")
    _cw(tbl.rows[1].cells[3], profil_inv, size=8, bg="F5F5F5")
    _cw(tbl.rows[1].cells[4], f"{score_str}/100 — {score_label}", bold=True, size=9, bg=risque_bg)

    doc.add_paragraph()

    # ── 2. Divergences
    _sub_heading(doc, "2.  Divergences identifiées")

    div_lower = str(divergence).lower()
    if div_lower not in ("aucune", "—", "", "none"):
        _key_bloc(doc,
                  "⚡  DIVERGENCE :",
                  f"{divergence}. "
                  "Ce signal dissonant doit être pris en compte avant toute décision — "
                  "les signaux techniques et fondamentaux ne convergent pas.",
                  "FFEB9C", "7D6608")
    else:
        p_nd = doc.add_paragraph()
        p_nd.paragraph_format.space_after = Pt(4)
        r_nd = p_nd.add_run("✔  Aucune divergence majeure identifiée. Cohérence technique/fondamental.")
        r_nd.font.size = Pt(9)
        r_nd.italic = True
        r_nd.font.color.rgb = _rgb("0F9D58")

    # ── 3. Recommandation finale
    _sub_heading(doc, "3.  Recommandation d'investissement")

    action_map = {
        "ACHAT FORT":  ("RENFORCER SIGNIFICATIVEMENT", "C6EFCE", "0F9D58",
                        "Tous les signaux sont alignés. C'est le moment d'augmenter l'exposition. "
                        "Priorité absolue dans l'allocation."),
        "ACHAT":       ("RENFORCER PROGRESSIVEMENT", "C6EFCE", "0F9D58",
                        "Signaux favorables. Entrée progressive recommandée sur 2-3 séances. "
                        "Ne pas entrer en totalité en une seule transaction."),
        "SURVEILLER":  ("CONSERVER ET SURVEILLER", "FFEB9C", "7D6608",
                        "Maintenir la position existante. Ne pas renforcer pour l'instant. "
                        "Attendre la prochaine confirmation directionnelle."),
        "PRUDENCE":    ("ALLÉGER PARTIELLEMENT", "FFF0E6", "C0392B",
                        "Réduire l'exposition de 25-50%. Sécuriser une partie des plus-values. "
                        "Conserver le solde en attente de signal de retournement."),
        "ÉVITER":      ("ALLÉGER OU SORTIR", "FFC7CE", "D93025",
                        "Signaux défavorables confirmés. Réduire ou clôturer la position. "
                        "Ne pas initier de nouvelle entrée sur ce titre."),
    }
    action_label, action_bg, action_fg, action_text = action_map.get(
        str(decision).upper(),
        ("SURVEILLER", "FFEB9C", "7D6608", "Maintenir la position. Aucune action urgente requise.")
    )

    # Grand bloc action
    tbl2 = doc.add_table(rows=1, cols=1)
    cell2 = tbl2.rows[0].cells[0]
    _cell_bg(cell2, action_bg)
    _cell_margins(cell2, 140)
    p2 = cell2.paragraphs[0]
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_action = p2.add_run(f"▶  {action_label}")
    r_action.bold = True
    r_action.font.size = Pt(13)
    r_action.font.color.rgb = _rgb(action_fg)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)

    # Texte d'accompagnement de la recommandation
    _narrative(doc, action_text)

    # Synthèse finale
    _narrative(doc,
               f"{ticker} — Score {score_str}/100 ({score_label}). "
               f"Reco : {reco}. Confiance : {confiance}. "
               f"Risque : {risque} — Horizon : {horizon.lower()}.")

    # Résumé compact final
    if resume and resume not in ("", "—"):
        _key_bloc(doc, "RÉSUMÉ :", resume, "E8F0FB", "1558A7")


# ── Pied de page ──────────────────────────────────────────────────────────────

def _pied(doc, date_str: str, freq: str = "JOUR", period_info: dict = None):
    p = doc.add_paragraph()
    _cp(p, 4, 0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "4")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), "CCCCCC")
    pBdr.append(top)
    pPr.append(pBdr)
    period_suffix = ""
    if freq != "JOUR" and period_info:
        period_suffix = (
            f"   |   {period_info.get('freq_label', freq)} : "
            f"{period_info.get('date_debut', '—')} → {period_info.get('date_fin', '—')}"
            f" ({period_info.get('nb_seances', '—')} séances)"
        )
    r = p.add_run(f"Document confidentiel — Analyse BRVM — {date_str}{period_suffix}")
    r.font.size = Pt(7)
    r.italic = True
    r.font.color.rgb = _rgb("999999")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ── Assemblage de la fiche ────────────────────────────────────────────────────

def _build_fiche_docx(s: dict, date_str: str, freq: str = "JOUR", period_info: dict = None) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    m = Cm(_MARGIN_CM)
    section.top_margin = m
    section.bottom_margin = m
    section.left_margin = m
    section.right_margin = m

    if doc.paragraphs:
        p0 = doc.paragraphs[0]
        p0.clear()
        _cp(p0, 0, 0)

    build_header(doc, s, date_str)           # En-tête : ticker, score, reco, indicateurs
    _add_separator(doc)
    build_market_table(doc, s)               # Métriques de marché
    _add_separator(doc)
    build_chart_comment(doc, s)              # Graphique + commentaire courbe
    _add_separator(doc)
    build_technical_analysis(doc, s)         # Analyse technique (tableau + convergence)
    _add_separator(doc)
    build_fundamental_analysis(doc, s)       # Analyse fondamentale + risques + perspectives
    _add_separator(doc)
    build_conclusion(doc, s)                 # Conclusion : matrice + divergences + action
    _pied(doc, date_str, freq, period_info)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Point d'entrée ────────────────────────────────────────────────────────────

def generate(docs_bytes, freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Génère une fiche Word par société depuis un ou plusieurs .docx source.
    Pipeline : texte → extraction LLM (minimal) → enrichissement Python → Word.
    docs_bytes : bytes (un seul doc) ou list[bytes] (plusieurs docs, plus récent en premier).
    Retourne list de (filename: str, docx_bytes: bytes).
    """
    if isinstance(docs_bytes, bytes):
        docs_bytes = [docs_bytes]

    date_str = date.today().strftime("%d/%m/%Y")
    date_file = date.today().strftime("%Y%m%d")
    freq_suffix = {"JOUR": "JOUR", "HEBDO": "HEBDO", "MENSUEL": "MENSUEL",
                   "TRIM": "TRIM", "ANNUEL": "ANNUEL"}.get(freq, freq)

    print(f"  [Fiches/{freq}] Étape 1/3 : Extraction du texte ({len(docs_bytes)} doc(s))...")
    full_text = _build_context(docs_bytes, freq)
    print(f"  [Fiches/{freq}] Texte source : {len(full_text)} chars")

    print(f"  [Fiches/{freq}] Étape 2/3 : Extraction LLM (JSON minimal)...")
    raw_companies = extract_all(full_text, freq, period_info)
    print(f"  [Fiches/{freq}] LLM → {len(raw_companies)} société(s) extraite(s).")

    if not raw_companies:
        print(f"  [Fiches/{freq}] AVERTISSEMENT : aucune société extraite, abandon.")
        return []

    print(f"  [Fiches/{freq}] Étape 3/3 : Enrichissement Python + génération Word...")
    all_companies = enrich(raw_companies)
    print(f"  [Fiches/{freq}] Enrichissement → {len(all_companies)} société(s).")

    results = []
    for company in all_companies:
        ticker = str(company.get("ticker") or "").strip()
        if not ticker:
            print(f"  [Fiches/{freq}] SKIP : société sans ticker")
            continue
        try:
            docx_bytes = _build_fiche_docx(company, date_str, freq, period_info)
            filename = f"Fiche_{ticker}_{date_file}_{freq_suffix}.docx"
            results.append((filename, docx_bytes))
            print(f"  [Fiches/{freq}] ✓ {filename}")
        except Exception as e:
            print(f"  [Fiches/{freq}] AVERTISSEMENT : fiche {ticker} ignorée — {e}")

    print(f"  [Fiches/{freq}] TOTAL : {len(results)} fiche(s) générée(s).")
    return results
