import io
import json
import os
from datetime import date, timedelta

import anthropic
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from dotenv import load_dotenv

load_dotenv()

_DESTINATAIRE = (
    "À l'attention de Madame la Directrice\n"
    "de l'Antenne Nationale de Bourse de Côte d'Ivoire"
)
_IA_MENTION = "Analyse multi-IA : DeepSeek · Gemini · Mistral"


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
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(12)
        run.font.color.rgb = _rgb("1A73E8")
    else:
        run.font.size = Pt(10)
        run.font.color.rgb = _rgb("333333")
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
    r3 = fp.add_run(f"{_IA_MENTION}    Page ")
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
    return "\n".join(lines)


# ── Contexte multi-documents ─────────────────────────────────────────────────

def _build_context(docs_bytes: list, freq: str) -> str:
    """Construit le texte source depuis une liste de documents (plus récent en premier)."""
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
            f"Dans tes champs textuels (analyses, contextes, implication, etc.), "
            f"décris l'ÉVOLUTION sur la période, pas seulement l'état actuel. "
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
  "capitalisation": "6 750 Mds FCFA",
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
  "secteurs": [{{"secteur": "Banque", "perf_100j": "+8.2%", "vs_brvm": "+3.0%", "position": "Surperformance", "lecture": "Solide"}}],
  "focus_industriels_telecoms": "analyse comparative",
  "implication_brvm": "concentration bancaire et diversification nécessaire",
  "top5_opportunites": [{{"symbole": "SGBCI", "score": 82, "raison": "raison"}}],
  "macro_international": "contexte mondial",
  "macro_africain": "contexte africain",
  "macro_uemoa": "contexte UEMOA",
  "impact_bourses_mondiales": "impact",
  "impact_indicateurs_brvm": "impact",
  "impact_societes_cotees": "impact",
  "sources_macro": [{{"source": "FMI", "resume": "résumé", "alerte": "vert"}}],
  "documents_officiels": [{{"type": "AG", "societe": "SGBCI", "date": "2026-04-30", "detail": "détail"}}],
  "alertes_google": [{{"mot_cle": "BRVM", "sentiment": "positif", "resume": "résumé", "url": "https://...", "score_pertinence": 8}}],
  "liquidite_haute": [{{"symbole": "SGBCI", "detail": "volume 50M FCFA/j"}}],
  "liquidite_risque": [{{"symbole": "XXXX", "detail": "faible volume"}}],
  "matrice_risques": [{{"symbole": "SGBCI", "risque": "faible", "horizon": "MT", "detail": "stable"}}],
  "divergences": [{{"symbole": "XXXX", "detail": "tech haussier / fond baissier"}}],
  "classement_47": [{{"symbole": "SGBCI", "secteur": "Banque", "prix": "14500", "score": 82, "signal_tech": "Haussier", "signal_fond": "Positif", "reco": "ACHAT"}}],
  "portefeuille_defensif": [{{"symbole": "SGBCI", "poids": 20}}, {{"symbole": "CASH", "poids": 15}}],
  "portefeuille_equilibre": [{{"symbole": "SGBCI", "poids": 25}}, {{"symbole": "CASH", "poids": 10}}],
  "portefeuille_offensif": [{{"symbole": "SGBCI", "poids": 30}}, {{"symbole": "CASH", "poids": 5}}],
  "rsi_surachat": [{{"symbole": "XXXX", "rsi": 75}}],
  "rsi_survente": [{{"symbole": "YYYY", "rsi": 25}}],
  "cours_bornes_100j": [{{"symbole": "XXXX", "detail": "proche borne haute"}}],
  "divergences_tech_fond": [{{"symbole": "XXXX", "reco": "NEUTRE", "detail": "RSI surachat + fondamentaux faibles"}}],
  "valeur_a_eviter": "XXXX",
  "recommandations_brvm": ["Améliorer la liquidité", "Diversifier avec nouvelles IPO", "Transparence des données", "Innovation IA"]
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
    """Retourne le meilleur score disponible ou 'N/D'."""
    for key in ("score", "score_composite", "score_technique"):
        v = item.get(key)
        if v is not None:
            return str(v)
    return "N/D"


def _get_score_float(item) -> float:
    """Retourne le score numérique pour le tri."""
    for key in ("score", "score_composite", "score_technique"):
        v = item.get(key)
        if v is not None:
            try:
                return float(v)
            except (ValueError, TypeError):
                pass
    return 0.0


# ── Pages ─────────────────────────────────────────────────────────────────────

def _page1(doc, d, date_str):
    # ── En-tête unique — apparaît une seule fois avant la première section ───
    p_hdr = doc.add_paragraph()
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(6)
    _bold(p_hdr, _DESTINATAIRE + "\n", 10)
    _normal(p_hdr, f"Date : {date_str}    |    {_IA_MENTION}", 9)

    _heading(doc, "NOTE DE SYNTHÈSE & ÉTAT DU MARCHÉ")

    # Bloc période (multi-périodes uniquement)
    pi = d.get("_period_info")
    if pi and pi.get("nb_seances", 1) > 1:
        p_pi = doc.add_paragraph()
        p_pi.paragraph_format.space_before = Pt(0)
        p_pi.paragraph_format.space_after = Pt(4)
        _bold(
            p_pi,
            f"📅 Période : {pi.get('date_debut', '—')} → {pi.get('date_fin', '—')}"
            f"   |   Séances analysées : {pi.get('nb_seances', '—')}"
            f"   |   Fréquence : {pi.get('freq_label', '—')}",
            9,
            "283593",
        )

    _heading(doc, "Résumé exécutif", 2)
    sentiment = _s(d, "sentiment")
    brvm_composite = _s(d, "brvm_composite")
    brvm_variation = _s(d, "brvm_variation")
    capitalisation = _s(d, "capitalisation")
    nb_achat = _s(d, "nb_achat", 0)
    nb_neutre = _s(d, "nb_neutre", 0)
    nb_vente = _s(d, "nb_vente", 0)
    top_opp = _s(d, "top_opportunite")
    principale_div = _s(d, "principale_divergence")
    perf_hist = _s(d, "perf_historique_100j")
    nb_div = len(_sl(d, "divergences"))

    narrative = (
        f"Le marché BRVM affiche un sentiment {sentiment}. "
        f"Le BRVM Composite s'établit à {brvm_composite} ({brvm_variation}) "
        f"avec une capitalisation de {capitalisation}. "
        f"On note {nb_achat} signaux d'ACHAT, {nb_neutre} NEUTRE et {nb_vente} VENTE. "
        f"La top opportunité du jour est {top_opp}. "
    )
    if nb_div:
        narrative += (
            f"{nb_div} sociétés présentent des divergences entre signaux "
            f"techniques et fondamentaux. "
        )
    if principale_div != "—":
        narrative += f"Principale divergence : {principale_div}. "
    narrative += f"La performance historique sur 100 jours est de {perf_hist}."

    p = doc.add_paragraph()
    _normal(p, narrative, 10)

    _heading(doc, "3 Risques majeurs", 2)
    for r in _sl(d, "risques")[:3]:
        doc.add_paragraph(f"• {r}", style="List Bullet")

    _heading(doc, "3 Opportunités majeures", 2)
    for o in _sl(d, "opportunites_majeures")[:3]:
        doc.add_paragraph(f"• {o}", style="List Bullet")


def _page2(doc, d):
    _heading(doc, "SYNTHÈSE GÉNÉRALE & ANALYSE SECTORIELLE")

    for label, key in [
        ("BRVM Composite sur 100 jours", "brvm_composite_100j"),
        ("Capitalisation", "capitalisation"),
    ]:
        p = doc.add_paragraph()
        _bold(p, f"{label} : ", 10)
        _normal(p, _s(d, key), 10)

    for title, key, header_bg, header_fg in [
        ("Top 10 Achats", "top10_achats", "0F9D58", "FFFFFF"),
        ("Flop 10", "flop10", "D93025", "FFFFFF"),
    ]:
        _heading(doc, title, 2)
        items = _sl(d, key)
        if items:
            tbl = doc.add_table(rows=1, cols=3)
            tbl.style = "Table Grid"
            _tbl_header(tbl, ["Symbole", "Score", "Reco"], header_bg, header_fg)
            for item in items[:10]:
                row = tbl.add_row()
                reco = str(item.get("reco", "NEUTRE"))
                row.cells[0].text = str(item.get("symbole", "—"))
                row.cells[1].text = _get_score(item)
                row.cells[2].text = reco
                _cell_bg(row.cells[2], _reco_color(reco))
        doc.add_paragraph()

    _heading(doc, "Tableau sectoriel", 2)
    secteurs = _sl(d, "secteurs")
    if secteurs:
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Secteur", "Perf.100j", "vs BRVM", "Position", "Lecture stratégique"])
        for s in secteurs:
            row = tbl.add_row()
            for i, k in enumerate(["secteur", "perf_100j", "vs_brvm", "position", "lecture"]):
                row.cells[i].text = str(s.get(k, "—"))
    doc.add_paragraph()

    for label, key in [
        ("Focus Industriels vs Télécoms", "focus_industriels_telecoms"),
        ("Implication BRVM", "implication_brvm"),
    ]:
        p = doc.add_paragraph()
        _bold(p, f"{label} : ", 10)
        _normal(p, _s(d, key), 10)

    _heading(doc, "Top 5 Opportunités — Score composite", 2)
    for op in _sl(d, "top5_opportunites")[:5]:
        p = doc.add_paragraph()
        _bold(p, f"{op.get('symbole', '—')} (score {op.get('score', '—')}) : ", 10)
        _normal(p, op.get("raison", "—"), 10)


def _page3(doc, d):
    _heading(doc, "ANALYSE MACRO INTERNATIONALE")

    for label, key in [
        ("Contexte international", "macro_international"),
        ("Contexte africain", "macro_africain"),
        ("Contexte UEMOA", "macro_uemoa"),
        ("Impact sur bourses mondiales", "impact_bourses_mondiales"),
        ("Impact sur indicateurs BRVM", "impact_indicateurs_brvm"),
        ("Impact sur sociétés cotées", "impact_societes_cotees"),
    ]:
        p = doc.add_paragraph()
        _bold(p, f"{label} : ", 10)
        _normal(p, _s(d, key, "Données non disponibles"), 10)

    _heading(doc, "Sources & Niveau d'alerte", 2)
    sources = _sl(d, "sources_macro")
    if sources:
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Source", "Résumé", "Alerte"])
        for s in sources:
            row = tbl.add_row()
            row.cells[0].text = str(s.get("source", "—"))
            row.cells[1].text = str(s.get("resume", "—"))
            alerte = str(s.get("alerte", "vert")).lower()
            emoji = "🟢" if alerte == "vert" else ("🟡" if alerte == "orange" else "🔴")
            row.cells[2].text = emoji
            _cell_bg(row.cells[2], "C6EFCE" if alerte == "vert" else ("FFEB9C" if alerte == "orange" else "FFC7CE"))


def _page4(doc, d):
    _heading(doc, "ACTUALITÉS DU MARCHÉ BRVM")

    _heading(doc, "A — Documents officiels (AG, dividendes, convocations, résultats)", 2)
    docs = _sl(d, "documents_officiels")
    if docs:
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Type", "Société", "Date", "Détail"], "1A73E8", "FFFFFF")
        for item in docs:
            row = tbl.add_row()
            row.cells[0].text = str(item.get("type") or "—")
            row.cells[1].text = str(item.get("societe") or "—")
            row.cells[2].text = str(item.get("date") or "—")
            row.cells[3].text = str(item.get("detail") or "—")
            _cell_bg(row.cells[0], "E8F0FE")
    else:
        doc.add_paragraph("Aucun document officiel recensé.")

    doc.add_paragraph()
    _heading(doc, "B — Alertes Google (groupées par mot-clé)", 2)
    alertes = _sl(d, "alertes_google")
    if alertes:
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Mot-clé", "Sentiment", "Résumé", "Source URL", "★/10"])
        for item in alertes:
            row = tbl.add_row()
            sent = str(item.get("sentiment", "neutre")).lower()
            emoji = "🟢" if sent == "positif" else ("🔴" if sent == "negatif" else "⚪")
            row.cells[0].text = str(item.get("mot_cle", "—"))
            row.cells[1].text = emoji
            row.cells[2].text = str(item.get("resume", "—"))
            row.cells[3].text = str(item.get("url", "—"))
            row.cells[4].text = str(item.get("score_pertinence", "—"))
    else:
        doc.add_paragraph("Aucune alerte Google recensée.")


def _page5(doc, d):
    _heading(doc, "MATRICE RISQUES & LIQUIDITÉ")

    _heading(doc, "Analyse liquidité", 2)
    _bold(doc.add_paragraph(), "Haute liquidité :", 10)
    for item in _sl(d, "liquidite_haute"):
        p = doc.add_paragraph()
        _bold(p, f"  {item.get('symbole', '—')} : ", 10)
        _normal(p, item.get("detail", "—"), 10)

    doc.add_paragraph()
    _bold(doc.add_paragraph(), "Titres à risque liquidité :", 10)
    for item in _sl(d, "liquidite_risque"):
        p = doc.add_paragraph()
        _bold(p, f"  {item.get('symbole', '—')} : ", 10)
        _normal(p, item.get("detail", "—"), 10)

    doc.add_paragraph()
    _heading(doc, "Matrice Risque / Horizon", 2)
    matrice = _sl(d, "matrice_risques")
    if matrice:
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        _tbl_header(tbl, ["Symbole", "Risque", "Horizon", "Commentaire"])
        for item in matrice:
            row = tbl.add_row()
            risque = str(item.get("risque", "moyen")).lower()
            row.cells[0].text = str(item.get("symbole", "—"))
            row.cells[1].text = risque.capitalize()
            _cell_bg(row.cells[1], "C6EFCE" if risque == "faible" else ("FFEB9C" if risque == "moyen" else "FFC7CE"))
            row.cells[2].text = str(item.get("horizon", "—"))
            row.cells[3].text = str(item.get("detail", "—"))

    doc.add_paragraph()
    _heading(doc, "Focus Divergences", 2)
    for item in _sl(d, "divergences"):
        p = doc.add_paragraph()
        _bold(p, f"⚠️ {item.get('symbole', '—')} : ", 10)
        _normal(p, item.get("detail", "—"), 10)


def _page6(doc, d):
    _heading(doc, "CLASSEMENT 47 SOCIÉTÉS /100")

    classement = _sl(d, "classement_47")
    if not classement:
        doc.add_paragraph("Données de classement non disponibles dans le rapport source.")
        return

    classement_sorted = sorted(classement, key=_get_score_float, reverse=True)

    tbl = doc.add_table(rows=1, cols=7)
    tbl.style = "Table Grid"
    _tbl_header(tbl, ["Symbole", "Secteur", "Prix", "Score /100", "Signal Tech", "Signal Fond", "Reco"], "1A73E8", "FFFFFF")

    for item in classement_sorted:
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


def _page7(doc, d):
    _heading(doc, "PORTEFEUILLES MODÈLES")

    configs = [
        ("🔵 Portefeuille Défensif", "portefeuille_defensif", "Faible risque · Haute liquidité · ~15% cash", "BDE9F7"),
        ("🟢 Portefeuille Équilibré", "portefeuille_equilibre", "Meilleur ratio rendement/risque · ~10% cash", "C6EFCE"),
        ("🔴 Portefeuille Offensif", "portefeuille_offensif", "Meilleurs scores composites · ~5% cash", "FFC7CE"),
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
            tbl.rows[0].cells[0].paragraphs[0].runs[0].bold = True
            tbl.rows[0].cells[1].paragraphs[0].runs[0].bold = True
            for item in items:
                row = tbl.add_row()
                row.cells[0].text = str(item.get("symbole", "—"))
                poids = item.get("poids")
                row.cells[1].text = f"{poids}%" if poids is not None else "N/D"
                if str(item.get("symbole", "")).upper() == "CASH":
                    _cell_bg(row.cells[0], "F5F5F5")
                    _cell_bg(row.cells[1], "F5F5F5")
        doc.add_paragraph()


def _page8(doc, d, date_str):
    _heading(doc, "ALERTES DU JOUR & RECOMMANDATIONS")

    for title, key, prefix in [
        ("RSI Surachat (> 70)", "rsi_surachat", "⚠️"),
        ("RSI Survente (< 30)", "rsi_survente", "📉"),
    ]:
        items = _sl(d, key)
        if items:
            _heading(doc, title, 2)
            for item in items:
                doc.add_paragraph(f"  {prefix} {item.get('symbole', '—')} — RSI : {item.get('rsi', '—')}")

    bornes = _sl(d, "cours_bornes_100j")
    if bornes:
        _heading(doc, "Cours proches des bornes 100 jours", 2)
        for item in bornes:
            p = doc.add_paragraph()
            _bold(p, f"  {item.get('symbole', '—')} : ", 10)
            _normal(p, item.get("detail", "—"), 10)

    div = _sl(d, "divergences_tech_fond")
    if div:
        _heading(doc, "Divergences Technique / Fondamental ⚠️", 2)
        for item in div:
            p = doc.add_paragraph()
            _bold(p, f"  {item.get('symbole', '—')} [{item.get('reco', '—')}] : ", 10)
            _normal(p, item.get("detail", "—"), 10)

    valeur_eviter = _s(d, "valeur_a_eviter")
    if valeur_eviter != "—":
        _heading(doc, "Valeur à éviter — Signal VENTE 🔴", 2)
        p = doc.add_paragraph()
        _bold(p, f"  {valeur_eviter}", 10, "D93025")

    _heading(doc, "Recommandations BRVM", 2)
    for r in _sl(d, "recommandations_brvm"):
        doc.add_paragraph(f"• {r}", style="List Bullet")

    doc.add_paragraph()
    _heading(doc, "Méthodologie Multi-IA & Précautions institutionnelles", 2)
    p = doc.add_paragraph(
        "Cette note est produite par un système multi-IA (DeepSeek · Gemini · Mistral) "
        "à titre indicatif. Les analyses ne constituent pas des conseils en investissement. "
        "Tout investissement comporte des risques de perte en capital. "
        "Usage strictement réservé aux professionnels habilités."
    )
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.color.rgb = _rgb("666666")
    p.runs[0].italic = True

    doc.add_paragraph()
    p = doc.add_paragraph()
    _bold(p, f"Note générée le {date_str}  —  Agent GitHub Rapports BRVM", 9)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


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

    _page1(doc, data, date_str)
    _page2(doc, data)
    _page3(doc, data)
    _page4(doc, data)
    _page5(doc, data)
    _page6(doc, data)
    _page7(doc, data)
    _page8(doc, data, date_str)

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
