import io
import json
import os
from datetime import date, datetime, timedelta

import anthropic
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from dotenv import load_dotenv

load_dotenv()

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


def _partie_heading(doc, text: str):
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = _rgb("1A73E8")
    return p


def _add_separator(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "CCCCCC")
    pBdr.append(bot)
    pPr.append(pBdr)


def _narrative(doc, text: str, size: int = 9, italic: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.italic = italic
    return p


# ── Helpers métier ────────────────────────────────────────────────────────────

def _signal_emoji(signal: str) -> str:
    s = str(signal or "").lower()
    if any(w in s for w in ("haussier", "positif", "achat", "fort", "bon", "élevé", "eleve")):
        return "🟢"
    if any(w in s for w in ("baissier", "négatif", "negatif", "vente", "faible")):
        return "🔴"
    return "🟡"


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


# ── Extraction via Claude ─────────────────────────────────────────────────────

def _get_tickers(full_text: str) -> list:
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=512,
        messages=[{
            "role": "user",
            "content": (
                "Extrais UNIQUEMENT la liste des tickers/symboles boursiers de toutes "
                "les sociétés BRVM présentes dans ce rapport.\n"
                "Retourne UNIQUEMENT ce JSON : {\"tickers\": [\"SGBCI\", \"SONATEL\", ...]}\n\n"
                f"RAPPORT :\n{full_text[:20000]}"
            ),
        }],
    )
    raw = msg.content[0].text.strip()
    start, end = raw.find("{"), raw.rfind("}") + 1
    if start == -1 or end == 0:
        return []
    try:
        return json.loads(raw[start:end]).get("tickers", [])
    except json.JSONDecodeError:
        return []


def _extract_batch(full_text: str, tickers: list, freq: str = "JOUR", period_info: dict = None) -> list:
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    _period_descs = {
        "HEBDO": "7 derniers jours (synthèse hebdomadaire)",
        "MENSUEL": "30 derniers jours (bilan mensuel)",
        "TRIM": "dernier trimestre (bilan trimestriel)",
        "ANNUEL": "dernière année (bilan annuel)",
    }
    if freq != "JOUR":
        nb = (period_info or {}).get("nb_seances", "?")
        period_ctx = (
            f"Tu analyses une synthèse sur les {_period_descs.get(freq, freq)} ({nb} séances). "
            f"Pour chaque société, décris l'ÉVOLUTION sur la période dans les champs textuels.\n"
        )
    else:
        period_ctx = ""

    prompt = f"""{period_ctx}Extrais les données de ces sociétés BRVM depuis le rapport : {', '.join(tickers)}

RAPPORT :
{full_text[:35000]}

Retourne UNIQUEMENT une liste JSON. Chaque société suit ce schéma (null si absent) :
{{
  "ticker": "SGBCI", "nom": "Société Générale de Banques en Côte d'Ivoire", "secteur": "Banque",
  "cours": "14500", "var_1j": "+0.5%",
  "score": 82, "reco": "ACHAT", "confiance": "Élevée",
  "mm": "haussier", "boll": "neutre", "macd": "haussier", "rsi": "neutre", "stoch": "haussier",
  "volatilite": "3.2%", "beta": "0.85", "liquidite": "haute",
  "risque": "faible", "divergence": "aucune", "stabilite": "bonne",
  "cours_debut_100j": "14000",
  "perf_100j": "+3.6%",
  "plus_haut_100j": "15200",
  "plus_bas_100j": "13500",
  "tendance_100j": "haussière",
  "volume_moyen_100j": "50M FCFA/j",
  "analyse_cours_100j": "Texte narratif fluide de 6-8 lignes max couvrant la performance sur 100j, le cours de départ vs actuel, les bornes haute/basse, la tendance générale et le volume moyen. Professionnel et sans liste à puces.",
  "mm_valeur": "14450",
  "mm_signal": "Haussier",
  "mm_detail": "La MM20 (14 200) reste au-dessus de la MM50 (13 900), confirmant un alignement haussier des moyennes.",
  "boll_sup": "15000",
  "boll_inf": "13800",
  "boll_signal": "Neutre",
  "boll_detail": "Le cours évolue dans le canal médian des bandes de Bollinger, sans signal de rupture imminent.",
  "macd_valeur": "0.15",
  "macd_signal_line": "0.08",
  "macd_signal": "Haussier",
  "macd_detail": "Le MACD (0.15) reste au-dessus de sa ligne de signal (0.08), indiquant un momentum positif persistant.",
  "rsi_valeur": "62",
  "rsi_signal": "Neutre",
  "rsi_detail": "Le RSI à 62 se situe en zone neutre, sans signal de surachat ni de survente.",
  "stoch_valeur": "75",
  "stoch_signal": "Surachat modéré",
  "stoch_detail": "Le Stochastique (75) approche la zone de surachat, à surveiller pour un potentiel retournement.",
  "synthese_tech": "Synthèse technique en 2-3 lignes : bilan global des 5 indicateurs et recommandation technique.",
  "analyse_fond": "Analyse fondamentale globale (max 80c)",
  "analyse_fond_recente": "Analyse fondamentale basée UNIQUEMENT sur les documents datant de moins de 18 mois. 4-6 lignes max. null si aucun document récent disponible.",
  "docs": [{{"type":"AG","date":"15/05/2026","impact":"positif"}}],
  "resume_rapport": "Performance solide (max 80c)",
  "indicateurs_fin": "PER 8x, DY 5% (max 60c)",
  "reco_src": "vert",
  "risques": ["risque 1", "risque 2"],
  "perspectives": "Positives à MT (max 60c)"
}}

FORMAT FINAL : [{{"ticker":"..."}}, ...]"""

    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    start, end = raw.find("["), raw.rfind("]") + 1
    if start == -1 or end == 0:
        return []
    try:
        return json.loads(raw[start:end])
    except json.JSONDecodeError:
        return []


# ── Sections de la fiche ──────────────────────────────────────────────────────

def _bandeau(doc, s: dict, date_str: str):
    ticker = _s(s, "ticker", "???")
    nom = _s(s, "nom", ticker)
    secteur = _s(s, "secteur", "—")
    score = s.get("score")
    score_label, score_color = _score_label_color(score)
    reco = _s(s, "reco", "NEUTRE")

    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"

    # Ligne 1 : [TICKER]  [NOM COMPLET]  [SECTEUR]  [DATE]
    r0 = tbl.rows[0]
    _cw(r0.cells[0], ticker, bold=True, size=13, color="FFFFFF", bg="1A237E")
    _cw(r0.cells[1], nom[:40], size=9, color="E8EAF6", bg="1A237E")
    _cw(r0.cells[2], secteur, size=8, color="C5CAE9", bg="283593")
    _cw(r0.cells[3], date_str, size=8, color="C5CAE9", bg="283593",
        align=WD_ALIGN_PARAGRAPH.RIGHT)

    # Ligne 2 : Score : X/100  [Appréciation]  |  RECOMMANDATION  |  🟢MM 🔴Boll ...
    r1 = tbl.rows[1]
    score_str = str(score) if score is not None else "—"
    _cw(r1.cells[0], f"Score : {score_str}/100  [{score_label}]",
        bold=True, size=9, color=score_color, bg="E8EAF6")
    _cw(r1.cells[1], reco, bold=True, size=9, color=_reco_fg(reco), bg=_reco_bg(reco))

    ind = "  ".join([
        f"{_signal_emoji(_s(s, 'mm'))}MM",
        f"{_signal_emoji(_s(s, 'boll'))}Boll",
        f"{_signal_emoji(_s(s, 'macd'))}MACD",
        f"{_signal_emoji(_s(s, 'rsi'))}RSI",
        f"{_signal_emoji(_s(s, 'stoch'))}Stoch",
    ])
    r1.cells[2].merge(r1.cells[3])
    _cw(r1.cells[2], ind, size=8, bg="E8EAF6")


def _bloc_metriques(doc, s: dict):
    _add_separator(doc)

    pairs = [
        ("Cours actuel",         _s(s, "cours", "—")),
        ("Variation 1j",         _s(s, "var_1j", "—")),
        ("Volatilité",           _s(s, "volatilite", "—")),
        ("Bêta",                 _s(s, "beta", "—")),
        ("Liquidité",            _s(s, "liquidite", "—")),
        ("Risque",               _s(s, "risque", "—")),
        ("Divergence tech/fond", _s(s, "divergence", "—")),
        ("Stabilité",            _s(s, "stabilite", "—")),
    ]

    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = "Table Grid"
    for i in range(4):
        ll, lv = pairs[i]
        rl, rv = pairs[i + 4]
        _cw(tbl.rows[i].cells[0], ll, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[1], lv, size=8)
        _cw(tbl.rows[i].cells[2], rl, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[3], rv, size=8)


def _placeholder_graphique(doc):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(
        "Courbe d'évolution du cours de l'action sur les 100 derniers jours\n"
        "avec prédictions"
    )
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("888888")


def _partie1_evolution(doc, s: dict):
    _add_separator(doc)
    _partie_heading(doc, "PARTIE 1 : ANALYSE DE L'ÉVOLUTION DU COURS (100 derniers jours)")

    analyse = _s(s, "analyse_cours_100j")
    if analyse and analyse not in ("", "—"):
        _narrative(doc, analyse)
        return

    ticker = _s(s, "ticker", "ce titre")
    cours = _s(s, "cours", "—")
    cours_debut = _s(s, "cours_debut_100j", "—")
    perf = _s(s, "perf_100j", "—")
    plus_haut = _s(s, "plus_haut_100j", "—")
    plus_bas = _s(s, "plus_bas_100j", "—")
    tendance = _s(s, "tendance_100j", "—")
    volume = _s(s, "volume_moyen_100j", "—")

    texte = (
        f"{ticker} a évolué de {cours_debut} FCFA à {cours} FCFA au cours des 100 dernières "
        f"séances, enregistrant une performance de {perf}. "
        f"Le titre a atteint un plus haut de {plus_haut} FCFA et un plus bas de {plus_bas} FCFA "
        f"sur la période, témoignant d'une amplitude notable. "
        f"La tendance générale est {tendance}. "
        f"Le volume moyen d'échanges s'établit à {volume}."
    )
    _narrative(doc, texte)


def _partie2_technique(doc, s: dict):
    _add_separator(doc)
    _partie_heading(doc, "PARTIE 2 : ANALYSE TECHNIQUE")

    indicateurs = [
        ("Moyennes Mobiles (MM)", "mm", "mm_valeur",   "mm_detail"),
        ("Bandes de Bollinger",  "boll", "boll_sup",   "boll_detail"),
        ("MACD",                 "macd", "macd_valeur", "macd_detail"),
        ("RSI",                  "rsi",  "rsi_valeur",  "rsi_detail"),
        ("Stochastique",         "stoch","stoch_valeur","stoch_detail"),
    ]

    for label, sig_key, val_key, detail_key in indicateurs:
        signal = _s(s, sig_key)
        valeur = _s(s, val_key)
        detail = _s(s, detail_key)
        emoji = _signal_emoji(signal)

        if detail and detail not in ("", "—"):
            ligne = f"{emoji} {label}"
            if valeur and valeur not in ("", "—"):
                ligne += f" ({valeur})"
            ligne += f" : {detail}"
        elif valeur and valeur not in ("", "—"):
            ligne = f"{emoji} {label} ({valeur}) — signal {signal}."
        else:
            ligne = f"{emoji} {label} — signal {signal}."

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(ligne)
        r.font.size = Pt(9)

    synthese = _s(s, "synthese_tech") or _s(s, "analyse_tech")
    if synthese and synthese not in ("", "—"):
        doc.add_paragraph()
        _narrative(doc, synthese, italic=True)


def _partie3_fondamentale(doc, s: dict):
    _add_separator(doc)
    _partie_heading(doc, "PARTIE 3 : ANALYSE FONDAMENTALE")

    analyse = _s(s, "analyse_fond_recente")
    if not analyse or analyse in ("", "—", "null", "None"):
        analyse = _s(s, "analyse_fond")

    if not analyse or analyse in ("", "—", "null", "None"):
        p = doc.add_paragraph("Aucune publication récente disponible.")
        p.paragraph_format.space_after = Pt(4)
        p.runs[0].font.size = Pt(9)
        p.runs[0].italic = True
    else:
        _narrative(doc, analyse)

    risques = _sl(s, "risques")
    if risques:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_after = Pt(2)
        r_r = p_r.add_run("⚠️ Risques : " + "   |   ".join(str(x) for x in risques[:2]))
        r_r.font.size = Pt(8)
        r_r.font.color.rgb = _rgb("E37400")

    perspectives = _s(s, "perspectives")
    if perspectives and perspectives not in ("", "—"):
        p_p = doc.add_paragraph()
        p_p.paragraph_format.space_after = Pt(2)
        r_p = p_p.add_run(f"🔮 Perspectives : {perspectives}")
        r_p.font.size = Pt(8)


def _pied(doc, date_str: str, freq: str = "JOUR", period_info: dict = None):
    p = doc.add_paragraph()
    _cp(p, 3, 0)
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
    r = p.add_run(f"Document confidentiel — {date_str}{period_suffix}")
    r.font.size = Pt(7)
    r.italic = True
    r.font.color.rgb = _rgb("999999")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ── Build d'une fiche ─────────────────────────────────────────────────────────

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

    _bandeau(doc, s, date_str)
    _bloc_metriques(doc, s)
    _placeholder_graphique(doc)
    _partie1_evolution(doc, s)
    _partie2_technique(doc, s)
    _partie3_fondamentale(doc, s)
    _pied(doc, date_str, freq, period_info)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Point d'entrée ────────────────────────────────────────────────────────────

def generate(docs_bytes, freq: str = "JOUR", period_info: dict = None) -> list:
    """
    Génère une fiche Word par société depuis un ou plusieurs .docx source.
    docs_bytes : bytes (un seul doc) ou list[bytes] (plusieurs docs, plus récent en premier).
    Retourne list de (filename: str, docx_bytes: bytes).
    """
    if isinstance(docs_bytes, bytes):
        docs_bytes = [docs_bytes]

    date_str = date.today().strftime("%d/%m/%Y")
    date_file = date.today().strftime("%Y%m%d")
    freq_suffix = {"JOUR": "JOUR", "HEBDO": "HEBDO", "MENSUEL": "MENSUEL", "TRIM": "TRIM", "ANNUEL": "ANNUEL"}.get(freq, freq)

    print(f"  [Fiches/{freq}] Extraction du texte source ({len(docs_bytes)} doc(s))...")
    full_text = _build_context(docs_bytes, freq)

    print(f"  [Fiches/{freq}] Récupération de la liste des tickers...")
    tickers = _get_tickers(full_text)
    if not tickers:
        print(f"  [Fiches/{freq}] AVERTISSEMENT : aucun ticker trouvé, abandon.")
        return []
    print(f"  [Fiches/{freq}] {len(tickers)} société(s) trouvée(s).")

    all_companies = []
    batch_size = 8
    for i in range(0, len(tickers), batch_size):
        batch = tickers[i:i + batch_size]
        batch_num = i // batch_size + 1
        total = (len(tickers) + batch_size - 1) // batch_size
        print(f"  [Fiches/{freq}] Extraction batch {batch_num}/{total} : {', '.join(batch)}")
        companies = _extract_batch(full_text, batch, freq, period_info)
        all_companies.extend(companies)

    results = []
    for company in all_companies:
        ticker = str(company.get("ticker") or "").strip()
        if not ticker:
            continue
        try:
            docx_bytes = _build_fiche_docx(company, date_str, freq, period_info)
            filename = f"Fiche_{ticker}_{date_file}_{freq_suffix}.docx"
            results.append((filename, docx_bytes))
        except Exception as e:
            print(f"  [Fiches/{freq}] AVERTISSEMENT : fiche {ticker} ignorée — {e}")

    print(f"  [Fiches/{freq}] {len(results)} fiche(s) générée(s).")
    return results
