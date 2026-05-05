import io
import json
import logging
import os
from datetime import date

logger = logging.getLogger(__name__)

import anthropic
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from dotenv import load_dotenv

load_dotenv()

_DESTINATAIRE = (
    "À l'attention de Madame Corinne Houmou ORMON\n"
    "Directrice de l'Antenne Nationale de Bourse de Côte d'Ivoire"
)


# ── Helpers docx ──────────────────────────────────────────────────────────────

def _cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _rgb(hex6: str) -> RGBColor:
    return RGBColor(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))


def _bold(paragraph, text: str, size: int = 11, color: str = None):
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = _rgb(color)
    return run


def _normal(paragraph, text: str, size: int = 10):
    run = paragraph.add_run(text)
    run.font.size = Pt(size)
    return run


def _heading(doc, text: str, level: int = 1):
    style_map = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}
    p = doc.add_paragraph(style=style_map.get(level, "Heading 1"))
    p.paragraph_format.space_before = Pt(10 if level == 1 else 6)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(14)
        run.font.color.rgb = _rgb("1A73E8")
    elif level == 2:
        run.font.size = Pt(11)
        run.font.color.rgb = _rgb("333333")
    else:
        run.font.size = Pt(10)
        run.font.color.rgb = _rgb("555555")
    return p


def _para(doc, text: str, size: int = 10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p


def _add_page_number(run):
    for tag, text in [("begin", None), (None, " PAGE "), ("end", None)]:
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
            run._r.append(el)
        else:
            el = OxmlElement("w:instrText")
            el.text = text
            el.set(qn("xml:space"), "preserve")
            run._r.append(el)


def _setup_header_footer(doc, date_str: str):
    section = doc.sections[0]

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.clear()
    hp.paragraph_format.space_after = Pt(2)
    r1 = hp.add_run(f"NOTE STRATÉGIQUE BRVM — {date_str}    ")
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = _rgb("1A73E8")
    r2 = hp.add_run("CONFIDENTIEL")
    r2.bold = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = _rgb("D93025")
    pPr = hp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "4")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "1A73E8")
    pBdr.append(bot)
    pPr.append(pBdr)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.clear()
    fp.paragraph_format.space_before = Pt(0)
    r3 = fp.add_run("Page ")
    r3.font.size = Pt(8)
    r3.font.color.rgb = _rgb("999999")
    pn = fp.add_run()
    _add_page_number(pn)
    pn.font.size = Pt(8)
    pn.font.color.rgb = _rgb("999999")


def _reco_color(reco: str) -> str:
    r = str(reco).upper()
    if "ACHAT" in r:
        return "C6EFCE"
    if "VENTE" in r:
        return "FFC7CE"
    return "FFEB9C"


def _tbl_header(tbl, headers, bg="E8F0FE", fg="333333"):
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]
        c.text = h
        _cell_bg(c, bg)
        run = c.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = _rgb(fg)


def _bloc(doc, label: str, texte: str, bg: str, fg: str = "1A1A1A"):
    """Bloc coloré POINT CLÉ / RISQUE / OPPORTUNITÉ inséré dans le document."""
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.rows[0].cells[0]
    _cell_bg(cell, bg)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side in ("top", "bottom", "left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "100")
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)
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


def _sep_synthese(doc):
    """Séparateur visuel '▌ SYNTHÈSE EXÉCUTIVE'."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("▌  SYNTHÈSE EXÉCUTIVE")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("1558A7")


def _sep_detail(doc):
    """Séparateur visuel '▌ ANALYSE DÉTAILLÉE'."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("▌  ANALYSE DÉTAILLÉE")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("666666")


def _graph_ph(doc, texte: str):
    """Placeholder de graphique centré en italique gris."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(texte)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("888888")


def _graph_lecture(doc, texte: str):
    """Lecture du graphique en italique discret sous le placeholder."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Lecture du graphique — " + texte)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("444444")


def _interp_tendance(tendance: str, perf: str) -> str:
    """Interprétation analytique approfondie d'une tendance de marché."""
    t = str(tendance or "").lower()
    positif = str(perf or "").strip().startswith("+")
    if "haussier" in t and positif:
        return (
            "Cette dynamique haussière traduit un regain d'appétit pour le risque, "
            "porté par une demande institutionnelle soutenue sur les valeurs de référence. "
            "La conjonction d'une tendance de fond positive et d'une performance cumulée favorable "
            "constitue un signal de continuité : les investisseurs peuvent envisager de maintenir, "
            "voire de renforcer leurs expositions sur les titres affichant les meilleurs scores composites. "
            "La sélectivité reste néanmoins de mise — les valeurs à faible liquidité doivent être traitées "
            "avec prudence même dans un contexte porteur."
        )
    if "haussier" in t:
        return (
            "La tendance de fond demeure haussière, même si la performance récente indique "
            "une phase de consolidation pouvant précéder une reprise du momentum. "
            "Ce type de configuration — tendance structurelle intacte mais progression momentanément freinée — "
            "représente souvent une opportunité d'entrée tactique pour les investisseurs n'ayant pas encore "
            "initié de position, sous réserve d'une confirmation technique à court terme. "
            "Il convient de surveiller les niveaux de support clés : un franchissement à la baisse "
            "remettrait en cause le scénario haussier de base."
        )
    if "baissier" in t:
        return (
            "Cette configuration baissière signale une prudence accrue des opérateurs, "
            "susceptible de refléter des prises de bénéfices massives ou des incertitudes macro-économiques "
            "pesant sur la valorisation des actifs cotés. "
            "Dans ce contexte, la priorité est à la préservation du capital : réduire les expositions "
            "sur les titres à faible score fondamental, privilégier le cash et les valeurs défensives "
            "à dividendes stables, et éviter tout renforcement sur momentum baissier. "
            "Un retournement ne pourra être confirmé qu'après une clôture franche au-dessus "
            "des résistances techniques identifiées."
        )
    return (
        "Le marché traverse une phase de neutralité et de consolidation, caractérisée par "
        "une absence de conviction directionnelle des opérateurs et des volumes en repli. "
        "Cette stabilité apparente peut masquer des tensions latentes — elle est souvent "
        "précurseur d'une rupture de tendance dont la direction doit être anticipée. "
        "La stratégie adaptée consiste à maintenir une allocation équilibrée, "
        "en réduisant l'exposition aux valeurs les plus volatiles et en conservant "
        "une réserve de liquidités permettant d'agir rapidement à la première confirmation directionnelle."
    )


def _lecture_graph_indice(idx: dict) -> str:
    """Analyse graphique approfondie (4 phrases) pour l'indice BRVM Composite."""
    niveau = idx["niveau"]
    haut = idx["plus_haut"]
    bas = idx["plus_bas"]
    perf = idx["perf_100j"]
    t = str(idx["tendance"]).lower()
    if "haussier" in t:
        return (
            f"La courbe décrit une progression structurée entre {bas} et {haut} points sur la période, "
            f"avec des phases de correction limitées qui ont systématiquement débouché sur de nouveaux plus hauts — "
            f"témoignant d'une structure technique fondamentalement saine. "
            f"Le gain cumulé de {perf} reflète un réel intérêt acheteur, non une simple effervescence spéculative. "
            f"À {niveau} points, l'indice se positionne en zone haute de son couloir de valorisation : "
            f"la résistance à {haut} constitue le prochain verrou à franchir pour ouvrir un nouveau palier haussier. "
            f"Un repli sous {bas} invaliderait le scénario haussier et inviterait à réviser l'allocation."
        )
    if "baissier" in t:
        return (
            f"L'évolution graphique révèle une pression vendeuse structurelle depuis {haut} vers {bas} points, "
            f"sans rebond technique significatif capable d'interrompre la séquence de plus bas successifs. "
            f"La performance de {perf} sur la période matérialise l'ampleur de la correction et "
            f"souligne l'absence d'acheteurs suffisamment convaincus pour absorber les flux vendeurs. "
            f"Le niveau actuel de {niveau} points représente un support critique : "
            f"son franchissement à la baisse ouvrirait la voie vers {bas} en première instance. "
            f"Toute tentative de rebond doit être confirmée en clôture avant d'envisager un repositionnement."
        )
    return (
        f"L'indice évolue dans un couloir de consolidation compris entre {bas} et {haut} points, "
        f"sans orientation directionnelle franche sur la période analysée. "
        f"La performance de {perf} illustre l'équilibre des forces en présence — "
        f"les acheteurs et les vendeurs s'annulant mutuellement sans qu'un camp ne prenne durablement l'avantage. "
        f"À {niveau} points, l'indice navigue en zone médiane de son range : "
        f"un bris au-dessus de {haut} serait le signal acheteur attendu, "
        f"un bris sous {bas} ouvrirait à la baisse. "
        f"La stratégie recommandée dans ce contexte : attente et préservation du capital."
    )


def _lecture_graph_capi(cap: dict) -> str:
    """Analyse graphique approfondie (4 phrases) pour la capitalisation boursière."""
    evolution = str(cap.get("evolution") or "").strip()
    valeur = cap["valeur"]
    pic = cap.get("pic", "—")
    plancher = cap.get("plancher", "—")
    if evolution.startswith("+"):
        pic_str = f", après avoir atteint un pic à {pic}" if pic != "—" else ""
        return (
            f"La progression de la capitalisation à {valeur}{pic_str} confirme une création de "
            f"valeur nette sur le marché — signe que les flux entrants surpassent les sorties de capitaux. "
            f"Cette dynamique positive reflète un intérêt institutionnel soutenu et une confiance "
            f"des opérateurs dans la solidité des fondamentaux des sociétés cotées. "
            f"Une capitalisation en hausse améliore mécaniquement la profondeur du marché et "
            f"réduit le risque de distorsion de prix sur les titres de référence. "
            f"Pour les investisseurs, ce contexte favorise les stratégies d'accumulation progressive "
            f"sur les valeurs à score composite élevé."
        )
    if evolution.startswith("-"):
        plancher_str = f", avec un plancher à {plancher} atteint sur la période" if plancher != "—" else ""
        return (
            f"Le recul de la capitalisation à {valeur}{plancher_str} signale une sortie nette de capitaux, "
            f"traduisant soit des prises de bénéfices coordonnées, soit une défiance des investisseurs "
            f"face à un contexte macro-économique ou sectoriel dégradé. "
            f"Une capitalisation en repli réduit la profondeur du marché et amplifie la sensibilité "
            f"du prix des titres aux ordres de vente — le risque de slippage est accru. "
            f"Cette configuration invite à une révision des pondérations : alléger les positions "
            f"sur les valeurs à faible liquidité en priorité, et renforcer la part de cash du portefeuille."
        )
    return (
        f"La capitalisation stabilisée à {valeur} traduit un équilibre entre flux acheteurs et vendeurs, "
        f"caractéristique d'une phase d'attente sur les marchés. "
        f"Cette stabilité peut être interprétée positivement — absence de panique vendeuse — "
        f"ou négativement — faiblesse des convictions acheteuses. "
        f"La profondeur de marché reste préservée, ce qui facilite les ajustements tactiques de portefeuille "
        f"sans impact majeur sur les prix. "
        f"La direction du prochain mouvement de capitalisation sera un indicateur avancé "
        f"de la tendance de l'indice dans les semaines à venir."
    )


# ── Extraction texte source ───────────────────────────────────────────────────

def _extract_text(doc_bytes: bytes) -> str:
    doc = Document(io.BytesIO(doc_bytes))
    lines = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(" | ".join(cells))
    text = "\n".join(lines)
    print(f"[DEBUG] Longueur texte extrait : {len(text)} chars")
    print(f"[DEBUG] Texte extrait (500 premiers chars) : {text[:500]}")
    return text


# ── Contexte multi-documents ─────────────────────────────────────────────────

def _build_context(docs_bytes: list, freq: str) -> str:
    if len(docs_bytes) == 1:
        return _extract_text(docs_bytes[0])

    max_older = {"HEBDO": 6, "MENSUEL": 9, "TRIM": 12, "ANNUEL": 14}.get(freq, 5)
    chars_older = {"HEBDO": 3000, "MENSUEL": 2000, "TRIM": 1500, "ANNUEL": 1000}.get(freq, 2000)

    recent = _extract_text(docs_bytes[0])[:25000]
    print(f"[DEBUG] Texte envoyé à Claude (200 premiers chars) : {recent[:200]}")
    older_parts = []
    for i, db in enumerate(docs_bytes[1:max_older + 1], 1):
        excerpt = _extract_text(db)[:chars_older]
        older_parts.append(f"--- Document J-{i} ---\n{excerpt}")

    return (
        "=== DOCUMENT LE PLUS RÉCENT (J) ===\n"
        + recent
        + ("\n\n=== HISTORIQUE (extraits) ===\n" + "\n\n".join(older_parts) if older_parts else "")
    )


# ── Extraction structurée via Claude ─────────────────────────────────────────

def _extract_data_with_claude(full_text: str, freq: str = "JOUR", period_info: dict = None) -> dict:
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    text = full_text[:40000]

    _period_descs = {
        "HEBDO": "7 derniers jours (synthèse hebdomadaire)",
        "MENSUEL": "30 derniers jours (bilan mensuel)",
        "TRIM": "dernier trimestre (bilan trimestriel)",
        "ANNUEL": "dernière année (bilan annuel complet)",
    }
    if freq != "JOUR":
        nb = (period_info or {}).get("nb_seances", "?")
        period_prefix = (
            f"Tu analyses une SYNTHÈSE sur les {_period_descs.get(freq, freq)} "
            f"couvrant {nb} séances. "
            f"Le document le plus récent est présenté en premier, suivi d'extraits historiques. "
            f"Dans tes champs textuels, décris l'ÉVOLUTION sur la période. "
            f"Pour les risques et opportunités, tiens compte des tendances observées.\n\n"
        )
    else:
        period_prefix = ""

    prompt = f"""Tu es un expert financier BRVM. {period_prefix}Extrais les données du rapport ci-dessous.
Retourne UNIQUEMENT un JSON valide, sans texte avant ou après. Utilise null si une donnée est absente.

RAPPORT SOURCE :
{text}

FORMAT JSON ATTENDU :
{{
  "sentiment": "Haussier|Baissier|Neutre",
  "brvm_composite": "245.67",
  "brvm_variation": "+0.45%",
  "brvm_composite_plus_haut_100j": "250.00",
  "brvm_composite_plus_bas_100j": "230.00",
  "brvm_composite_tendance": "haussière|baissière|latérale",
  "capitalisation": "6 750 Mds FCFA",
  "capitalisation_evolution_100j": "+3.2%",
  "capitalisation_pic_100j": "7 000 Mds FCFA",
  "capitalisation_plancher_100j": "6 400 Mds FCFA",
  "nb_achat": 12,
  "nb_neutre": 20,
  "nb_vente": 15,
  "top_opportunite": "SGBCI",
  "principale_divergence": "description",
  "perf_historique_100j": "+5.2%",
  "risques": ["risque 1", "risque 2", "risque 3"],
  "opportunites_majeures": ["opp 1", "opp 2", "opp 3"],
  "brvm_composite_100j": "de 230 à 245",
  "top10_achats": [{{"symbole": "SGBCI", "score": 82, "reco": "ACHAT"}}],
  "flop10": [{{"symbole": "XXXX", "score": 20, "reco": "VENTE"}}],
  "secteurs": [{{
    "secteur": "Banque",
    "nb_societes": 10,
    "perf_100j": "+8.2%",
    "vs_brvm": "+3.0%",
    "position": "Surperformance",
    "sentiment": "Positif",
    "risque": "Moyen",
    "prix_moyen": "12500",
    "societes": ["SGBCI", "BICICI", "BICIA-CI"]
  }}],
  "focus_industriels_telecoms": "analyse comparative",
  "implication_brvm": "concentration bancaire et diversification nécessaire",
  "top5_opportunites": [{{"symbole": "SGBCI", "score": 82, "raison": "raison"}}],
  "macro_international": "contexte mondial",
  "macro_africain": "contexte africain",
  "macro_uemoa": "contexte UEMOA",
  "impact_bourses_mondiales": "impact",
  "impact_indicateurs_brvm": "impact",
  "impact_societes_cotees": "impact",
  "synthese_macro": "synthèse narrative du contexte macro",
  "recommandation_macro": "recommandation finale macro",
  "sources_macro": [{{"source": "FMI", "resume": "résumé", "alerte": "vert"}}],
  "documents_officiels": [{{"type": "AG", "societe": "SGBCI", "date": "2026-04-30", "detail": "détail"}}],
  "alertes_google": [{{"mot_cle": "BRVM", "sentiment": "positif", "resume": "résumé", "score_pertinence": 8}}],
  "liquidite_haute": [{{
    "symbole": "SGBCI",
    "volume_moyen": "50M FCFA/j",
    "valeur_moyenne": "45M FCFA/j",
    "perf_100j": "+5%",
    "recommandation": "ACHAT",
    "detail": "volume 50M FCFA/j"
  }}],
  "liquidite_risque": [{{
    "symbole": "XXXX",
    "volume_moyen": "2M FCFA/j",
    "valeur_moyenne": "1.5M FCFA/j",
    "perf_100j": "-3%",
    "recommandation": "VENTE",
    "detail": "faible volume"
  }}],
  "matrice_risques": [{{"symbole": "SGBCI", "risque": "faible", "horizon": "MT", "detail": "stable"}}],
  "divergences": [{{"symbole": "XXXX", "detail": "tech haussier / fond baissier"}}],
  "classement_47": [{{"symbole": "SGBCI", "secteur": "Banque", "prix": "14500", "score": 82, "signal_tech": "Haussier", "signal_fond": "Positif", "reco": "ACHAT"}}],
  "portefeuille_defensif": [{{"symbole": "SGBCI", "poids": 20}}, {{"symbole": "CASH", "poids": 15}}],
  "portefeuille_equilibre": [{{"symbole": "SGBCI", "poids": 25}}, {{"symbole": "CASH", "poids": 10}}],
  "portefeuille_offensif": [{{"symbole": "SGBCI", "poids": 30}}, {{"symbole": "CASH", "poids": 5}}],
  "rsi_surachat": [{{
    "symbole": "XXXX",
    "rsi": 75,
    "cours_100j_max": "5490",
    "signal_tech": "Vente",
    "signal_fond": "Achat",
    "reco": "ACHAT",
    "detail": "RSI surachat, cours proche borne haute"
  }}],
  "rsi_survente": [{{
    "symbole": "YYYY",
    "rsi": 25,
    "cours_100j_min": "1200",
    "signal_tech": "Achat",
    "signal_fond": "Vente",
    "reco": "VENTE",
    "detail": "RSI survente, cours proche borne basse"
  }}],
  "cours_bornes_100j": [{{"symbole": "XXXX", "detail": "proche borne haute"}}],
  "divergences_tech_fond": [{{"symbole": "XXXX", "reco": "NEUTRE", "detail": "RSI surachat + fondamentaux faibles"}}],
  "valeur_a_eviter": "XXXX",
  "recommandations_brvm": ["Améliorer la liquidité", "Diversifier avec nouvelles IPO", "Transparence des données"],
  "conclusion_sentiment": "Haussier modéré avec prudence sectorielle",
  "top3_opportunites_conclusion": [{{"symbole": "SGBCI", "raison": "raison courte justifiée"}}],
  "top3_risques_conclusion": [{{"risque": "Risque de liquidité", "detail": "plusieurs titres peu liquides"}}],
  "recommandation_finale": "Le marché présente..."
}}"""

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text.strip()
    start = raw.find("{")
    end = raw.rfind("}") + 1
    if start == -1 or end == 0:
        return {}
    try:
        return json.loads(raw[start:end])
    except json.JSONDecodeError:
        return {}


# ── Helpers sûrs ─────────────────────────────────────────────────────────────

def _s(data, key, default="—"):
    v = data.get(key)
    return v if v is not None else default


def _sl(data, key):
    v = data.get(key)
    return v if isinstance(v, list) else []


def _get_score(item) -> str:
    for key in ("score", "score_composite", "score_technique"):
        v = item.get(key)
        if v is not None:
            return str(v)
    return "N/D"


def _get_score_float(item) -> float:
    for key in ("score", "score_composite", "score_technique"):
        v = item.get(key)
        if v is not None:
            try:
                return float(v)
            except (ValueError, TypeError):
                pass
    return 0.0


def _no_data(text) -> str:
    """Remplace 'Données insuffisantes' et variantes par une formulation neutre."""
    s = str(text).strip() if text else ""
    if s in ("", "—", "null", "None"):
        return "Aucune information significative n'a été collectée sur ce point."
    lower = s.lower()
    if any(x in lower for x in ("données insuffisantes", "données non disponibles", "non disponible")):
        return "Aucune information significative n'a été collectée sur ce point."
    return s


# ── Fonctions d'extraction et de synthèse ────────────────────────────────────

def extract_brvm_index_data(data: dict) -> dict:
    """Extrait les champs de l'indice BRVM Composite depuis les données structurées."""
    return {
        "niveau":    _s(data, "brvm_composite"),
        "variation": _s(data, "brvm_variation"),
        "perf_100j": _s(data, "perf_historique_100j"),
        "plus_haut": _s(data, "brvm_composite_plus_haut_100j"),
        "plus_bas":  _s(data, "brvm_composite_plus_bas_100j"),
        "tendance":  _s(data, "brvm_composite_tendance"),
        "range":     _s(data, "brvm_composite_100j"),
    }


def extract_market_cap_data(data: dict) -> dict:
    """Extrait les champs de la capitalisation boursière depuis les données structurées."""
    return {
        "valeur":    _s(data, "capitalisation"),
        "evolution": _s(data, "capitalisation_evolution_100j"),
        "pic":       _s(data, "capitalisation_pic_100j"),
        "plancher":  _s(data, "capitalisation_plancher_100j"),
    }


def summarize_section(items: list, fmt_fn, max_items: int = 5) -> list:
    """Retourne les max_items premiers éléments formatés par fmt_fn."""
    return [fmt_fn(item) for item in items[:max_items] if item]


# ── Section 1 — En-tête institutionnel ───────────────────────────────────────

def _section_entete(doc, d, date_str: str):
    """Section 1 — En-tête institutionnel."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    r1 = p.add_run(_DESTINATAIRE + "\n")
    r1.bold = True
    r1.font.size = Pt(11)
    r2 = p.add_run(f"Date : {date_str}")
    r2.font.size = Pt(10)

    pi = d.get("_period_info")
    if pi and pi.get("nb_seances", 1) > 1:
        p_pi = doc.add_paragraph()
        p_pi.paragraph_format.space_after = Pt(6)
        _bold(
            p_pi,
            f"Période : {pi.get('date_debut', '—')} → {pi.get('date_fin', '—')}"
            f"   |   Séances : {pi.get('nb_seances', '—')}"
            f"   |   Fréquence : {pi.get('freq_label', '—')}",
            9, "283593",
        )


def _section_brvm_composite(doc, d):
    """Section 2 — Indice BRVM Composite sur 100 jours."""
    _heading(doc, "ÉVOLUTION DE L'INDICE BRVM COMPOSITE (100 JOURS)")

    # ── Données préalables
    idx = extract_brvm_index_data(d)
    nb_achat = _s(d, "nb_achat", "—")
    nb_neutre = _s(d, "nb_neutre", "—")
    nb_vente = _s(d, "nb_vente", "—")
    sentiment = _s(d, "sentiment", "Neutre")
    t = str(idx["tendance"]).lower()
    top_opp = _s(d, "top_opportunite", None)

    # ── Blocs exécutifs (lecture prioritaire)
    _bloc(
        doc, "POINT CLÉ :",
        f"Indice à {idx['niveau']} pts | Tendance : {idx['tendance']} | Perf 100j : {idx['perf_100j']} | "
        f"Sentiment : {sentiment} | {nb_achat} signaux ACHAT — {nb_vente} signaux VENTE.",
        "E8F0FB", "1558A7",
    )
    if "haussier" in t:
        opp_msg = f"{top_opp} identifiée comme valeur prioritaire. " if top_opp and top_opp != "—" else ""
        _bloc(doc, "OPPORTUNITÉ :",
              f"{opp_msg}Marché haussier : maintenir les positions. "
              f"Renforcer les titres à score ≥ 70/100 au-dessus de leur MM50.",
              "E8F8F0", "155724")
    elif "baissier" in t:
        _bloc(doc, "RISQUE :",
              "Tendance baissière confirmée. "
              "Réduire les expositions. Renforcer le cash. "
              "Ne pas moyenner à la baisse.",
              "FFF0E6", "C0392B")
    else:
        _bloc(doc, "RISQUE :",
              "Marché en consolidation. Pas de signal directionnel franc. "
              "Attendre la confirmation avant toute prise de position.",
              "FFEB9C", "7D6608")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    _para(doc,
          f"L'indice clôture à {idx['niveau']} points ({idx['variation']} sur la séance). "
          f"Performance sur 100 jours : {idx['perf_100j']}. "
          f"Amplitude : {idx['plus_bas']} → {idx['plus_haut']} points. "
          + (f"Range observé : {idx['range']}." if idx["range"] != "—" else ""))
    _para(doc,
          f"Signaux agrégés : {nb_achat} ACHAT — {nb_neutre} NEUTRE — {nb_vente} VENTE. "
          f"Sentiment dominant : {sentiment}. "
          f"Cette répartition est cohérente avec la tendance {idx['tendance']} de l'indice.")

    # ── Analyse détaillée
    _sep_detail(doc)
    _para(doc, _interp_tendance(idx["tendance"], idx["perf_100j"]))

    # ── Graphique
    _graph_ph(doc, "[Courbe d'évolution de l'indice BRVM Composite sur les 100 derniers jours avec prédictions]")
    _graph_lecture(doc, _lecture_graph_indice(idx))


def _section_capitalisation(doc, d):
    """Section 3 — Capitalisation boursière globale."""
    _heading(doc, "ÉVOLUTION DE LA CAPITALISATION GLOBALE")

    # ── Données préalables
    cap = extract_market_cap_data(d)
    evolution = str(cap.get("evolution") or "").strip()

    # ── Blocs exécutifs
    if evolution.startswith("+"):
        _bloc(doc, "OPPORTUNITÉ :",
              f"Capitalisation en hausse de {evolution}. "
              f"Marché en création de valeur nette. "
              f"Contexte favorable à l'accumulation sur les valeurs à score élevé.",
              "E8F8F0", "155724")
    elif evolution.startswith("-"):
        _bloc(doc, "RISQUE :",
              f"Capitalisation en recul de {evolution}. "
              f"Sortie nette de capitaux. "
              f"Réduire le dimensionnement des positions. Privilégier les titres liquides.",
              "FFF0E6", "C0392B")
    else:
        _bloc(doc, "POINT CLÉ :",
              f"Capitalisation stable à {cap['valeur']}. "
              f"Équilibre acheteurs/vendeurs. "
              f"Attendre le prochain signal directionnel avant toute décision d'allocation.",
              "E8F0FB", "1558A7")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    _para(doc, f"Capitalisation totale BRVM : {cap['valeur']}.")
    if evolution != "—":
        parts = [f"Évolution 100j : {evolution}."]
        if cap["pic"] != "—":
            parts.append(f"Pic : {cap['pic']}.")
        if cap["plancher"] != "—":
            parts.append(f"Plancher : {cap['plancher']}.")
        _para(doc, "  ".join(parts))

    # ── Analyse détaillée
    _sep_detail(doc)
    if evolution.startswith("+"):
        _para(doc,
              "La hausse de la capitalisation traduit une création de richesse nette pour les actionnaires. "
              "Elle améliore la profondeur du marché. "
              "Un marché plus capitalisé est moins sujet aux distorsions de prix. "
              "Les conditions d'exécution s'améliorent pour les institutionnels. "
              "Ce signal est cohérent avec la tendance haussière de l'indice composite.")
    elif evolution.startswith("-"):
        _para(doc,
              "Le recul de la capitalisation reflète une pression vendeuse dominante. "
              "La profondeur de marché se réduit. "
              "L'impact prix des transactions importantes s'accentue. "
              "Ce contexte exige une révision des seuils de risque et une gestion active de la liquidité.")
    else:
        _para(doc,
              "La stabilité de la capitalisation indique un équilibre des flux. "
              "La profondeur de marché est préservée. "
              "L'attentisme des opérateurs suggère une attente de signal directionnel. "
              "La direction du prochain mouvement sera un indicateur avancé de tendance.")

    # ── Graphique
    _graph_ph(doc, "[Courbe d'évolution de la capitalisation boursière sur les 100 derniers jours]")
    _graph_lecture(doc, _lecture_graph_capi(cap))
    doc.add_paragraph()


# ── Section 4 — Analyse sectorielle ──────────────────────────────────────────

def _section_secteurs(doc, d):
    """Section 4 — Analyse sectorielle."""
    _heading(doc, "ANALYSE PAR SECTEUR")
    secteurs = _sl(d, "secteurs")
    if not secteurs:
        _para(doc, "Données sectorielles non disponibles dans le rapport source.")
        return

    # ── Calculs préalables
    def _perf_float(s):
        try:
            return float(str(s.get("perf_100j", "0")).replace("%", "").replace("+", "").replace(",", "."))
        except (ValueError, TypeError):
            return 0.0

    tries = sorted(secteurs, key=_perf_float, reverse=True)
    leader = tries[0] if tries else None
    laggard = tries[-1] if len(tries) > 1 else None
    nb_total = len(secteurs)
    sentiments = [str(s.get("sentiment", "")) for s in secteurs if s.get("sentiment")]
    nb_pos = sum(1 for s in sentiments if "positif" in s.lower() or "favorable" in s.lower())

    # ── Blocs exécutifs
    if leader:
        _bloc(doc, "OPPORTUNITÉ :",
              f"Secteur leader : {leader.get('secteur', '—')} ({leader.get('perf_100j', '—')}, "
              f"{leader.get('vs_brvm', '—')} vs BRVM). "
              f"Momentum acheteur concentré ici. Renforcer les meilleures valeurs du secteur.",
              "E8F8F0", "155724")
    if laggard and laggard.get("secteur") != (leader or {}).get("secteur"):
        _bloc(doc, "RISQUE :",
              f"Secteur en difficulté : {laggard.get('secteur', '—')} ({laggard.get('perf_100j', '—')}). "
              f"Ne pas moyenner à la baisse. Réallouer vers les secteurs en surperformance.",
              "FFF0E6", "C0392B")
    _bloc(doc, "POINT CLÉ :",
          f"{nb_pos}/{nb_total} secteurs affichent un sentiment positif. "
          f"La sélection sectorielle est le principal levier de surperformance dans ce contexte.",
          "E8F0FB", "1558A7")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    _para(doc,
          f"{nb_total} secteurs couverts. "
          f"{nb_pos} affichent un sentiment positif. "
          + (f"Leader : {leader.get('secteur', '—')} à {leader.get('perf_100j', '—')}. " if leader else "")
          + (f"Laggard : {laggard.get('secteur', '—')} à {laggard.get('perf_100j', '—')}." if laggard and laggard.get("secteur") != (leader or {}).get("secteur") else ""))

    # ── Tableau de synthèse sectorielle
    if len(secteurs) > 1:
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Secteur", "Perf 100j", "vs BRVM", "Sentiment", "Risque"], "1A73E8", "FFFFFF")
        for s in tries:
            row = tbl.add_row()
            row.cells[0].text = str(s.get("secteur") or "—")
            row.cells[1].text = str(s.get("perf_100j") or "—")
            row.cells[2].text = str(s.get("vs_brvm") or "—")
            row.cells[3].text = str(s.get("sentiment") or "—")
            row.cells[4].text = str(s.get("risque") or "—")
            pf = _perf_float(s)
            _cell_bg(row.cells[1], "C6EFCE" if pf > 0 else ("FFC7CE" if pf < 0 else "FFEB9C"))
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        doc.add_paragraph()

    # ── Analyse détaillée
    _sep_detail(doc)

    if leader:
        leader_socs = ", ".join((leader.get("societes") or [])[:4]) or "—"
        _heading(doc, f"Secteur leader : {leader.get('secteur', '—')}", 2)
        _para(doc,
              f"Performance 100j : {leader.get('perf_100j', '—')} — soit {leader.get('vs_brvm', '—')} vs BRVM Composite. "
              f"Sentiment : {leader.get('sentiment', '—')}. Risque : {leader.get('risque', '—')}. "
              f"{leader.get('nb_societes', '—')} sociétés cotées, dont : {leader_socs}.")
        _para(doc,
              "Ce secteur est le moteur principal de la progression du marché. "
              "Il capte une part disproportionnée des flux acheteurs institutionnels. "
              "Toute révision d'allocation doit y prioriser les valeurs à score composite ≥ 70/100.")

    if laggard and laggard.get("secteur") != (leader or {}).get("secteur"):
        _heading(doc, f"Secteur en difficulté : {laggard.get('secteur', '—')}", 2)
        _para(doc,
              f"Performance 100j : {laggard.get('perf_100j', '—')}. "
              f"Risque : {laggard.get('risque', '—')}.")
        _para(doc,
              "Sous-performance pouvant résulter de facteurs structurels (pression réglementaire, érosion des marges) "
              "ou conjoncturels (choc macro, déficit de liquidité). "
              "Procéder à une réévaluation rigoureuse des positions existantes. "
              "Biais de réduction progressive si les signaux ne s'améliorent pas sous 30 jours.")

    macro_uemoa = _s(d, "macro_uemoa")
    if macro_uemoa and macro_uemoa != "—":
        _para(doc,
              f"Lien macro : cette dynamique sectorielle s'inscrit dans un contexte UEMOA marqué par "
              f"{str(macro_uemoa)[:180]}. "
              "(Voir section Analyse Macro pour le détail complet.)")

    top5 = _sl(d, "top5_opportunites")
    if top5:
        _para(doc, "Opportunités cross-sectorielles : "
              + ", ".join(f"{x.get('symbole', '—')} — {x.get('raison', '—')[:50]}" for x in top5[:3])
              + ".")


# ── Section 5 — Analyse de liquidité ─────────────────────────────────────────

def _section_liquidite(doc, d):
    """Section 5 — Analyse de liquidité."""
    _heading(doc, "ANALYSE DE LIQUIDITÉ")

    # ── Données préalables
    haute = _sl(d, "liquidite_haute")
    risque_l = _sl(d, "liquidite_risque")
    matrice = _sl(d, "matrice_risques")

    # ── Blocs exécutifs
    if haute:
        syms_haute = ", ".join(item.get("symbole", "—") for item in haute[:4])
        _bloc(doc, "OPPORTUNITÉ :",
              f"Titres à haute liquidité : {syms_haute}. "
              f"Candidats naturels pour les prises de position de taille. "
              f"Priorité : ceux combinant liquidité élevée + score ≥ 70/100.",
              "E8F8F0", "155724")
    if risque_l:
        sym_risque = ", ".join(item.get("symbole", "—") for item in risque_l[:3])
        _bloc(doc, "RISQUE :",
              f"Liquidité insuffisante sur : {sym_risque}. "
              f"Risque de slippage et de blocage. "
              f"Plafond recommandé : 2-3% du portefeuille par titre illiquide.",
              "FFF0E6", "C0392B")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    _para(doc,
          f"{len(haute)} titre(s) à haute liquidité identifiés. "
          f"{len(risque_l)} titre(s) à liquidité risquée. "
          "La liquidité conditionne directement la flexibilité tactique du portefeuille.")
    _para(doc,
          "Trois risques opérationnels à surveiller : "
          "(1) slippage — écart entre prix théorique et prix d'exécution réel ; "
          "(2) blocage — impossibilité de sortir rapidement en cas de retournement ; "
          "(3) distorsion de prix — les volumes faibles amplifient l'impact des ordres importants.")

    # ── Analyse détaillée
    _sep_detail(doc)

    if haute:
        _heading(doc, "Titres à haute liquidité — opportunités opérationnelles", 2)
        _para(doc,
              "Ces valeurs permettent entrées et sorties sans impact prix significatif. "
              "Elles constituent le socle des portefeuilles institutionnels. "
              "Flexibilité tactique optimale pour la gestion active.")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Symbole", "Volume moyen / jour", "Perf 100j", "Recommandation"], "1A73E8", "FFFFFF")
        for item in haute[:6]:
            row = tbl.add_row()
            row.cells[0].text = str(item.get("symbole") or "—")
            row.cells[1].text = str(item.get("volume_moyen") or item.get("valeur_moyenne") or item.get("detail") or "—")
            row.cells[2].text = str(item.get("perf_100j") or "—")
            reco = str(item.get("recommandation") or "—")
            row.cells[3].text = reco
            _cell_bg(row.cells[3], _reco_color(reco))
            perf_str = str(item.get("perf_100j") or "")
            _cell_bg(row.cells[2], "C6EFCE" if perf_str.startswith("+") else ("FFC7CE" if perf_str.startswith("-") else "FFEB9C"))
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        doc.add_paragraph()
    else:
        _para(doc, "Aucun titre à haute liquidité identifié dans le rapport source.")

    if risque_l:
        _heading(doc, "Titres à liquidité risquée — contraintes opérationnelles", 2)
        _para(doc,
              "Volumes insuffisants pour absorber des ordres institutionnels sans distorsion de prix. "
              "Toute position doit être fractionnée sur plusieurs séances. "
              "Dimensionnement extrêmement prudent. "
              "Le coût de sortie doit être intégré dès l'entrée en position.")
        tbl2 = doc.add_table(rows=1, cols=4)
        tbl2.style = "Table Grid"
        _tbl_header(tbl2, ["Symbole", "Volume moyen / jour", "Perf 100j", "Recommandation"], "C0392B", "FFFFFF")
        for item in risque_l[:6]:
            row = tbl2.add_row()
            row.cells[0].text = str(item.get("symbole") or "—")
            row.cells[1].text = str(item.get("volume_moyen") or item.get("valeur_moyenne") or item.get("detail") or "—")
            row.cells[2].text = str(item.get("perf_100j") or "—")
            reco = str(item.get("recommandation") or "—")
            row.cells[3].text = reco
            _cell_bg(row.cells[3], _reco_color(reco))
            _cell_bg(row.cells[0], "FFF0E6")
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        doc.add_paragraph()

    if _sl(d, "secteurs") and haute:
        _para(doc,
              "Corrélation liquidité/secteur : les titres les plus liquides appartiennent généralement "
              "aux secteurs les plus capitalisés. "
              "Surpondérer les secteurs leaders cumule double avantage : "
              "meilleures perspectives de performance ET flexibilité opérationnelle supérieure. "
              "(Cf. Analyse par Secteur pour les détails.)")

    if matrice:
        _para(doc, "Matrice de risque synthétique : "
              + " | ".join(
                  f"{m.get('symbole', '—')} [{m.get('risque', '—')} / horizon {m.get('horizon', '—')}]"
                  for m in matrice[:5]) + ".")


# ── Section 6 — Analyse macro ────────────────────────────────────────────────

def _section_macro(doc, d):
    """Section 6 — Contexte macro international, africain & UEMOA."""
    _heading(doc, "ANALYSE MACRO — CONTEXTE INTERNATIONAL, AFRICAIN & UEMOA")

    # ── Données préalables
    def _val(key, max_chars=300):
        v = _s(d, key)
        if not v or v in ("—", "null", "None"):
            return None
        v = v.replace("★", "").strip()
        return v[:max_chars].rstrip() + ("…" if len(v) > max_chars else "")

    reco_macro = _val("recommandation_macro", 250)
    risques_macro = _sl(d, "risques")
    opps_macro = _sl(d, "opportunites_majeures")
    secteurs = _sl(d, "secteurs")
    leader_nom = "—"
    if secteurs:
        def _pf(s):
            try:
                return float(str(s.get("perf_100j", "0")).replace("%", "").replace("+", "").replace(",", "."))
            except (ValueError, TypeError):
                return 0.0
        leader_nom = sorted(secteurs, key=_pf, reverse=True)[0].get("secteur", "—")

    # ── Blocs exécutifs
    _bloc(doc, "POINT CLÉ :",
          (reco_macro or "Surveiller l'évolution du contexte macro pour ajuster l'exposition.")
          + f" Secteur le plus exposé : {leader_nom}.",
          "E8F0FB", "1558A7")
    if risques_macro:
        _bloc(doc, "RISQUE :",
              "Risques identifiés : " + " — ".join(str(r)[:80] for r in risques_macro[:3]) + ".",
              "FFF0E6", "C0392B")
    if opps_macro:
        _bloc(doc, "OPPORTUNITÉ :",
              "Opportunités identifiées : " + " — ".join(str(o)[:80] for o in opps_macro[:3]) + ".",
              "E8F8F0", "155724")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    parts_synth = []
    ctx_intl_s = _val("macro_international", 120)
    if ctx_intl_s:
        parts_synth.append(f"International : {ctx_intl_s}")
    ctx_afr_s = _val("macro_africain", 120)
    if ctx_afr_s:
        parts_synth.append(f"Afrique : {ctx_afr_s}")
    ctx_uemoa_s = _val("macro_uemoa", 120)
    if ctx_uemoa_s:
        parts_synth.append(f"UEMOA : {ctx_uemoa_s}")
    if parts_synth:
        _para(doc, "  |  ".join(parts_synth) + ".")
    imp_synth = _val("impact_indicateurs_brvm", 180)
    if imp_synth:
        _para(doc, f"Impact BRVM : {imp_synth}")

    # ── Analyse détaillée
    _sep_detail(doc)
    _para(doc,
          "La BRVM n'est pas imperméable aux chocs macro-économiques globaux et africains. "
          "Les flux de capitaux étrangers, les politiques monétaires de la BCEAO "
          "et les dynamiques des autres places africaines influencent directement la valorisation des actifs. "
          "L'analyse s'articule en trois niveaux imbriqués.")

    ctx_intl = _val("macro_international")
    if ctx_intl:
        _heading(doc, "Environnement international", 2)
        _para(doc, ctx_intl)
        impact_brvm = _val("impact_bourses_mondiales", 250)
        if impact_brvm:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _bold(p, "Implication pour la BRVM : ", 10, "1558A7")
            _normal(p, impact_brvm, 10)

    ctx_afr = _val("macro_africain")
    if ctx_afr:
        _heading(doc, "Environnement africain et régional", 2)
        _para(doc, ctx_afr)

    ctx_uemoa = _val("macro_uemoa")
    if ctx_uemoa:
        _heading(doc, "Zone UEMOA — facteurs de proximité", 2)
        _para(doc, ctx_uemoa)

    imp_indic = _val("impact_indicateurs_brvm", 250)
    imp_soc = _val("impact_societes_cotees", 250)
    if imp_indic or imp_soc:
        _heading(doc, "Impacts directs sur la BRVM", 2)
        if imp_indic:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _bold(p, "Sur les indicateurs de marché : ", 10)
            _normal(p, imp_indic, 10)
        if imp_soc:
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(4)
            p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _bold(p2, "Sur les sociétés cotées : ", 10)
            _normal(p2, imp_soc, 10)

    sources = _sl(d, "sources_macro")
    if sources:
        _para(doc, "Sources : "
              + " | ".join(f"{src.get('source', '—')} — {str(src.get('resume') or '—')[:80]}"
                           for src in sources[:4]) + ".")

    synthese = _val("synthese_macro", 350)
    if synthese:
        _heading(doc, "Synthèse macro et positionnement", 2)
        _para(doc, synthese)

    if secteurs and leader_nom != "—":
        _para(doc,
              f"Lien sectoriel : le secteur {leader_nom} (leader de performance) "
              f"est directement exposé aux conditions macro décrites ci-dessus. "
              "Calibrer l'exposition en intégrant simultanément signaux techniques, "
              "fondamentaux sectoriels et contexte macro.")


# ── Section 7 — Actualités du marché BRVM ────────────────────────────────────

def _section_actualites(doc, d):
    """Section 7 — Actualités du marché BRVM."""
    _heading(doc, "ACTUALITÉS DU MARCHÉ BRVM")

    # ── Données préalables
    docs_off = _sl(d, "documents_officiels")
    alertes = _sl(d, "alertes_google")
    top_opp = _s(d, "top_opportunite")
    div = _s(d, "principale_divergence")

    # ── Blocs exécutifs
    if top_opp and top_opp != "—":
        _bloc(doc, "OPPORTUNITÉ :",
              f"{top_opp} : valeur prioritaire sur signaux croisés (technique + fondamental + marché). "
              f"Signal d'achat conditionnel — confirmer sur 2 séances avant renforcement.",
              "E8F8F0", "155724")
    if div and div != "—":
        _bloc(doc, "POINT CLÉ :",
              f"Divergence identifiée : {str(div)[:180]}. "
              "À intégrer dans toute décision d'allocation.",
              "E8F0FB", "1558A7")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    nb_docs = len(docs_off)
    nb_alertes = len(alertes)
    _para(doc,
          f"{nb_docs} événement(s) corporate recensé(s). "
          f"{nb_alertes} signal(s) de veille presse. "
          + (f"Valeur focus : {top_opp}." if top_opp and top_opp != "—" else ""))

    # ── Analyse détaillée
    _sep_detail(doc)

    if docs_off:
        _heading(doc, "Événements corporate et documents officiels", 2)
        _para(doc,
              "Les événements corporate (AG, résultats, dividendes, avertissements) sont des catalyseurs directs. "
              "Ils peuvent déclencher des mouvements de cours significatifs à court terme.")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Type", "Société", "Date", "Détail / Impact"], "1A73E8", "FFFFFF")
        for item in docs_off[:8]:
            row = tbl.add_row()
            row.cells[0].text = str(item.get("type") or "Document").capitalize()
            row.cells[1].text = str(item.get("societe") or "—")
            row.cells[2].text = str(item.get("date") or "—")
            row.cells[3].text = str(item.get("detail") or "—")[:120]
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        doc.add_paragraph()
    else:
        _para(doc, "Aucun événement corporate significatif recensé pour cette période.")

    if alertes:
        _heading(doc, "Veille presse et signaux d'actualité", 2)
        _para(doc,
              "Signaux d'information classés par sentiment et score de pertinence. "
              "Ces signaux influencent la perception des investisseurs à court terme.")
        for item in alertes[:5]:
            sent = str(item.get("sentiment", "")).lower()
            score = item.get("score_pertinence", "—")
            mot_cle = str(item.get("mot_cle") or "—")
            resume = str(item.get("resume") or "—")[:200]
            prefixe = "POSITIF" if sent == "positif" else ("NÉGATIF" if "neg" in sent else "NEUTRE")
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r_label = p.add_run(f"[{prefixe} — {mot_cle} | Pertinence : {score}/10]  ")
            r_label.bold = True
            r_label.font.size = Pt(9)
            r_txt = p.add_run(resume)
            r_txt.font.size = Pt(9)
        doc.add_paragraph()

    if not docs_off and not alertes:
        _para(doc, "Aucune actualité significative recensée pour cette période.")

    if top_opp and top_opp != "—":
        _heading(doc, "Valeur focus", 2)
        _para(doc,
              f"{top_opp} cumule des signaux croisés favorables. "
              "Technique, fondamental et contexte de marché convergent. "
              "Surveiller les catalyseurs imminents (résultats, annonce, volume inhabituel). "
              "Valider le scénario sur 2 séances consécutives avant tout renforcement.")


# ── Section 10 — Alertes stratégiques ────────────────────────────────────────

def _section_alertes(doc, d):
    """Section 10 — Alertes stratégiques."""
    _heading(doc, "ALERTES STRATÉGIQUES DU JOUR")

    # ── Données préalables
    surachat = _sl(d, "rsi_surachat")
    survente = _sl(d, "rsi_survente")
    divergences = _sl(d, "divergences_tech_fond")
    cours_bornes = _sl(d, "cours_bornes_100j")
    valeur_eviter = _s(d, "valeur_a_eviter")
    nb_surachat = len(surachat)
    nb_survente = len(survente)
    nb_div = len(divergences)
    principale_div = _s(d, "principale_divergence")

    # ── Blocs exécutifs
    if nb_surachat > nb_survente and nb_surachat > 0:
        syms_sa = ", ".join(item.get("symbole", "—") for item in surachat[:3])
        _bloc(doc, "RISQUE :",
              f"{nb_surachat} signal(s) de surachat ({syms_sa}). "
              "Surextension technique. Ne pas renforcer. Sécuriser les plus-values.",
              "FFF0E6", "C0392B")
    if nb_survente > 0:
        syms_sv = ", ".join(item.get("symbole", "—") for item in survente[:3])
        _bloc(doc, "OPPORTUNITÉ :",
              f"Survente sur {syms_sv}. "
              "Points d'entrée potentiels si fondamentaux solides. "
              "Confirmer sur 2 séances avant action.",
              "E8F8F0", "155724")
    if principale_div and principale_div != "—":
        _bloc(doc, "POINT CLÉ :",
              f"Divergence principale : {str(principale_div)[:180]}. "
              "Signal structurel majeur — à peser dans toute décision d'allocation.",
              "E8F0FB", "1558A7")

    # ── Synthèse exécutive
    _sep_synthese(doc)
    total_alertes = nb_surachat + nb_survente + nb_div + len(cours_bornes) + (1 if valeur_eviter and valeur_eviter != "—" else 0)
    _para(doc,
          f"{total_alertes} alerte(s) identifiée(s) : "
          f"{nb_surachat} surachat — {nb_survente} survente — {nb_div} divergence(s).")
    if nb_surachat > nb_survente:
        _para(doc, "Profil dominant : surachat. Risque de correction à court terme accru. Sécuriser les positions.")
    elif nb_survente > nb_surachat:
        _para(doc, "Profil dominant : survente. Opportunités potentielles sur les valeurs à bons fondamentaux.")
    else:
        _para(doc, "Alertes équilibrées. Gestion au cas par cas — pas de biais directionnel dominant.")

    # ── Tableau des alertes
    _sep_detail(doc)
    _para(doc,
          "Chaque alerte requiert une vérification complémentaire avant exécution. "
          "Un signal RSI n'invalide pas les fondamentaux — croiser avec le score composite.")

    # ── Construction du tableau
    rows = []

    for item in surachat[:3]:
        sym = item.get("symbole", "—")
        rsi = item.get("rsi", "—")
        cours_max = item.get("cours_100j_max", "—")
        sig_tech = item.get("signal_tech", "—")
        sig_fond = item.get("signal_fond", "—")
        reco = item.get("reco", "—")
        detail = str(item.get("detail") or "")[:100]
        titre = f"{sym}  RSI : {rsi} | Cours max 100j : {cours_max}"
        if detail:
            titre += f" | {detail}"
        action = (
            f"Ne pas renforcer. Envisager une réduction partielle. "
            f"Signal tech : {sig_tech} / Signal fond : {sig_fond}. Reco : {reco}."
        )
        rows.append(("⚠  Surachat RSI", titre, action, "FFC7CE"))

    for item in survente[:3]:
        sym = item.get("symbole", "—")
        rsi = item.get("rsi", "—")
        cours_min = item.get("cours_100j_min", "—")
        sig_tech = item.get("signal_tech", "—")
        sig_fond = item.get("signal_fond", "—")
        reco = item.get("reco", "—")
        detail = str(item.get("detail") or "")[:100]
        titre = f"{sym}  RSI : {rsi} | Cours min 100j : {cours_min}"
        if detail:
            titre += f" | {detail}"
        action = (
            f"Surveiller le signal de rebond — entrée progressive si confirmation. "
            f"Signal tech : {sig_tech} / Signal fond : {sig_fond}. Reco : {reco}."
        )
        rows.append(("✔  Survente RSI", titre, action, "C6EFCE"))

    for item in divergences[:2]:
        sym = item.get("symbole", "—")
        reco = item.get("reco", "—")
        detail = str(item.get("detail") or "—")[:120]
        titre = f"{sym}  {detail}"
        action = (
            f"Approche mixte requise : ne pas agir sur le seul signal technique. "
            f"Attendre convergence technique + fondamental. Reco : {reco}."
        )
        rows.append(("↕  Divergence Tech/Fond", titre, action, "FFEB9C"))

    for item in cours_bornes[:1]:
        sym = item.get("symbole", "—")
        detail = str(item.get("detail") or "—")[:120]
        titre = f"{sym}  {detail}"
        rows.append(("📊  Cours en borne", titre, "Surveiller le franchissement — point de bascule imminent.", "E8F0FB"))

    if valeur_eviter and valeur_eviter != "—":
        rows.append((
            "✖  Valeur à éviter",
            f"{valeur_eviter} — Signaux cumulés négatifs",
            "Sortir ou ne pas initier de position. Risque global élevé.",
            "FFC7CE",
        ))

    # ── Tableau des alertes
    if rows:
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Type d'alerte", "Valeur / Signal", "Action recommandée"], "C0392B", "FFFFFF")
        for signal, titre, action, bg in rows:
            row = tbl.add_row()
            row.cells[0].text = signal
            row.cells[1].text = titre
            row.cells[2].text = action
            _cell_bg(row.cells[0], bg)
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
        doc.add_paragraph()
    else:
        _para(doc, "Aucune alerte stratégique majeure identifiée pour cette séance.")

    # ── Recommandation de lecture croisée
    if nb_surachat + nb_survente + nb_div > 0:
        _para(doc,
              "Recommandation transversale : croiser systématiquement chaque signal d'alerte "
              "avec le score composite (Classement 47 sociétés) avant toute décision. "
              "Un RSI en surachat sur un titre à score 80/100 et fondamentaux solides "
              "a un profil de risque très différent d'un titre à score 30/100.")


# ── Classement 47 sociétés ────────────────────────────────────────────────────

def _section_classement(doc, d):
    _heading(doc, "CLASSEMENT 47 SOCIÉTÉS /100")

    classement = _sl(d, "classement_47")
    if not classement:
        _para(doc, "Données de classement non disponibles dans le rapport source.")
        return

    tbl = doc.add_table(rows=1, cols=7)
    tbl.style = "Table Grid"
    _tbl_header(tbl, ["Symbole", "Secteur", "Prix", "Score /100", "Signal Tech", "Signal Fond", "Reco"], "1A73E8", "FFFFFF")

    for item in sorted(classement, key=_get_score_float, reverse=True):
        row = tbl.add_row()
        score = _get_score_float(item)
        reco = str(item.get("reco", "NEUTRE")).upper()
        values = [
            str(item.get("symbole") or "—"),
            str(item.get("secteur") or "—"),
            str(item.get("prix") or "—"),
            _get_score(item),
            str(item.get("signal_tech") or "—"),
            str(item.get("signal_fond") or "—"),
            reco,
        ]
        for i, v in enumerate(values):
            row.cells[i].text = v
            for p in row.cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8)
        _cell_bg(row.cells[3], "C6EFCE" if score >= 70 else ("FFEB9C" if score >= 50 else "FFC7CE"))
        _cell_bg(row.cells[6], _reco_color(reco))


# ── Portefeuilles modèles ─────────────────────────────────────────────────────

def _section_portefeuilles(doc, d):
    _heading(doc, "PORTEFEUILLES MODÈLES")

    configs = [
        ("Portefeuille Défensif", "portefeuille_defensif", "Faible risque · Haute liquidité · ~15% cash", "BDE9F7"),
        ("Portefeuille Équilibré", "portefeuille_equilibre", "Meilleur ratio rendement/risque · ~10% cash", "C6EFCE"),
        ("Portefeuille Offensif", "portefeuille_offensif", "Meilleurs scores composites · ~5% cash", "FFC7CE"),
    ]

    for label, key, desc, bg in configs:
        _heading(doc, label, 2)
        p = doc.add_paragraph(desc)
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.color.rgb = _rgb("666666")
        items = _sl(d, key)
        if items:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            _tbl_header(tbl, ["Symbole / Position", "Poids %"], bg)
            for item in items:
                row = tbl.add_row()
                row.cells[0].text = str(item.get("symbole", "—"))
                poids = item.get("poids")
                row.cells[1].text = f"{poids}%" if poids is not None else "N/D"
                if str(item.get("symbole", "")).upper() == "CASH":
                    _cell_bg(row.cells[0], "F5F5F5")
                    _cell_bg(row.cells[1], "F5F5F5")
        doc.add_paragraph()


# ── Build complet ─────────────────────────────────────────────────────────────

def _build_docx(data: dict, date_str: str) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    _setup_header_footer(doc, date_str)

    _section_entete(doc, data, date_str)       # 1 — En-tête institutionnel
    _section_brvm_composite(doc, data)         # 2 — Indice BRVM Composite (100j)
    _section_capitalisation(doc, data)         # 3 — Capitalisation globale
    _section_secteurs(doc, data)               # 4 — Analyse sectorielle
    _section_liquidite(doc, data)              # 5 — Analyse de liquidité
    _section_macro(doc, data)                  # 6 — Contexte macro + impacts BRVM
    _section_actualites(doc, data)             # 7 — Actualités du marché
    _section_classement(doc, data)             # 8 — Classement sociétés /100
    _section_portefeuilles(doc, data)          # 9 — Portefeuilles modèles
    _section_alertes(doc, data)                # 10 — Alertes du jour

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Point d'entrée ────────────────────────────────────────────────────────────

def generate(docs_bytes, freq: str = "JOUR", period_info: dict = None) -> tuple:
    """
    Génère la Note Stratégique BRVM depuis un ou plusieurs .docx source.
    docs_bytes : bytes (un seul doc) ou list[bytes] (plusieurs docs, plus récent en premier).
    Retourne (filename: str, docx_bytes: bytes).
    """
    if isinstance(docs_bytes, bytes):
        docs_bytes = [docs_bytes]

    date_str = date.today().strftime("%d/%m/%Y")
    date_file = date.today().strftime("%Y%m%d")

    freq_suffix = {"JOUR": "JOUR", "HEBDO": "HEBDO", "MENSUEL": "MENSUEL", "TRIM": "TRIM", "ANNUEL": "ANNUEL"}.get(freq, freq)

    print(f"  [Note/{freq}] Extraction du texte source ({len(docs_bytes)} doc(s))...")
    full_text = _build_context(docs_bytes, freq)

    print(f"  [Note/{freq}] Extraction des données structurées via Claude...")
    data = _extract_data_with_claude(full_text, freq, period_info or {})

    if period_info:
        data["_period_info"] = period_info
        data["_freq"] = freq

    print(f"  [Note/{freq}] Construction du document Word...")
    docx_bytes = _build_docx(data, date_str)

    filename = f"Note_Strategique_BRVM_{date_file}_{freq_suffix}.docx"
    return filename, docx_bytes
