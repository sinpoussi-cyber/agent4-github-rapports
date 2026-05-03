import io
import json
import logging
import os
from collections import defaultdict
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
    "À l'attention de Madame la Directrice\n"
    "de l'Antenne Nationale de Bourse de Côte d'Ivoire"
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


# ── Section 1 — En-tête + Synthèse générale ──────────────────────────────────

def _section_entete_synthese(doc, d, date_str: str):
    p_hdr = doc.add_paragraph()
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(10)
    run_dest = p_hdr.add_run(_DESTINATAIRE + "\n")
    run_dest.bold = True
    run_dest.font.size = Pt(11)
    run_date = p_hdr.add_run(f"Date : {date_str}")
    run_date.font.size = Pt(10)

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

    _heading(doc, "SYNTHÈSE GÉNÉRALE")

    brvm = _s(d, "brvm_composite")
    variation = _s(d, "brvm_variation")
    perf = _s(d, "perf_historique_100j")
    plus_haut = _s(d, "brvm_composite_plus_haut_100j")
    plus_bas = _s(d, "brvm_composite_plus_bas_100j")
    tendance = _s(d, "brvm_composite_tendance")
    range_100j = _s(d, "brvm_composite_100j")

    texte_composite = (
        f"Le BRVM Composite s'établit à {brvm} points, enregistrant une variation journalière de {variation}. "
        f"Sur les 100 derniers jours, l'indice a progressé de {perf}, évoluant entre un plus bas "
        f"de {plus_bas} et un plus haut de {plus_haut} points. "
        f"La tendance générale est {tendance}."
    )
    if range_100j != "—":
        texte_composite += f" L'amplitude observée sur la période est : {range_100j}."
    _para(doc, texte_composite)

    p_ph1 = doc.add_paragraph()
    p_ph1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_ph1.add_run("[Courbe d'évolution de l'indice composite sur les 100 derniers jours]")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = _rgb("888888")

    doc.add_paragraph()

    capi = _s(d, "capitalisation")
    capi_evol = _s(d, "capitalisation_evolution_100j")
    capi_pic = _s(d, "capitalisation_pic_100j")
    capi_plancher = _s(d, "capitalisation_plancher_100j")

    texte_capi = f"La capitalisation boursière globale s'élève à {capi}."
    if capi_evol != "—":
        texte_capi += f" Sur les 100 derniers jours, elle a évolué de {capi_evol}"
        if capi_pic != "—" and capi_plancher != "—":
            texte_capi += f", avec un pic à {capi_pic} et un plancher à {capi_plancher}."
        elif capi_pic != "—":
            texte_capi += f", avec un pic à {capi_pic}."
        elif capi_plancher != "—":
            texte_capi += f", avec un plancher à {capi_plancher}."
        else:
            texte_capi += "."
    _para(doc, texte_capi)

    p_ph2 = doc.add_paragraph()
    p_ph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p_ph2.add_run("[Courbe d'évolution de la capitalisation sur les 100 derniers jours]")
    r2.italic = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = _rgb("888888")


# ── Section 2 — Analyse par secteur ──────────────────────────────────────────

def _section_secteurs(doc, d):
    _heading(doc, "ANALYSE PAR SECTEUR")

    _para(doc,
        "Cette section présente une analyse comparative de tous les secteurs représentés "
        "à la BRVM, incluant la performance moyenne, le sentiment général du marché et "
        "le niveau de risque moyen."
    )

    secteurs = _sl(d, "secteurs")
    if not secteurs:
        _para(doc, "Données sectorielles non disponibles.")
        return

    def _perf_float(s):
        try:
            return float(str(s.get("perf_100j", "0")).replace("%", "").replace("+", "").replace(",", "."))
        except (ValueError, TypeError):
            return 0.0

    for s in sorted(secteurs, key=_perf_float, reverse=True):
        nom = s.get("secteur", "—")
        nb = s.get("nb_societes", "—")
        perf = s.get("perf_100j", "—")
        vs = s.get("vs_brvm", "—")
        position = s.get("position", "—")
        sentiment = s.get("sentiment", "—")
        risque = s.get("risque", "—")
        prix = s.get("prix_moyen", "—")
        societes = s.get("societes", [])

        indicateur = "🟢 Surperformance" if "Surperformance" in str(position) else "🔴 Sous-performance"
        societes_str = ", ".join(str(x) for x in societes) if societes else "—"

        texte = (
            f"Le secteur {nom} regroupe {nb} société(s) avec une performance moyenne de {perf} "
            f"sur 100 jours, soit {vs} par rapport au BRVM Composite. "
            f"Le sentiment général est {sentiment} avec un risque moyen {risque} "
            f"et un prix moyen de {prix} FCFA. {indicateur}. "
            f"Sociétés concernées : {societes_str}."
        )
        _heading(doc, nom, 2)
        _para(doc, texte)


# ── Section 3 — Analyse de liquidité ─────────────────────────────────────────

def _section_liquidite(doc, d):
    _heading(doc, "ANALYSE DE LIQUIDITÉ")

    _para(doc,
        "L'analyse de liquidité classe les titres selon leur accessibilité sur le marché. "
        "Les titres à haute liquidité offrent des conditions d'entrée et de sortie favorables, "
        "tandis que les titres à faible liquidité présentent un risque de sortie significatif "
        "pouvant affecter les conditions d'exécution des ordres."
    )

    _heading(doc, "Titres à Haute Liquidité", 2)
    haute = _sl(d, "liquidite_haute")
    if haute:
        for item in haute:
            sym = item.get("symbole", "—")
            vol = item.get("volume_moyen") or item.get("detail", "—")
            val = item.get("valeur_moyenne", "")
            perf = item.get("perf_100j", "")
            reco = item.get("recommandation", "")

            parts = [f"{sym} affiche un volume moyen de {vol}"]
            if val:
                parts.append(f"pour une valeur moyenne de {val}")
            if perf:
                parts.append(f"une performance sur 100 jours de {perf}")
            if reco:
                parts.append(f"recommandation : {reco}")
            _para(doc, ". ".join(parts) + ".")
    else:
        _para(doc, "Aucun titre à haute liquidité identifié.")

    _heading(doc, "Titres à Faible Liquidité — Risque Élevé", 2)
    risque = _sl(d, "liquidite_risque")
    if risque:
        for item in risque:
            sym = item.get("symbole", "—")
            vol = item.get("volume_moyen") or item.get("detail", "—")
            val = item.get("valeur_moyenne", "")
            perf = item.get("perf_100j", "")
            reco = item.get("recommandation", "")

            parts = [f"{sym} présente un volume moyen limité à {vol}"]
            if val:
                parts.append(f"valeur moyenne de {val}")
            if perf:
                parts.append(f"performance sur 100 jours de {perf}")
            if reco:
                parts.append(f"recommandation : {reco}")
            _para(doc, ". ".join(parts) + ".")
    else:
        _para(doc, "Aucun titre à risque liquidité identifié.")

    _para(doc,
        "Avertissement : les titres à faible liquidité peuvent présenter des écarts "
        "importants entre les prix d'achat et de vente, et les ordres de sortie peuvent "
        "nécessiter plusieurs séances pour être exécutés. Une attention particulière est "
        "recommandée lors de toute prise de position sur ces valeurs."
    )


# ── Section 4 — Analyse macro ─────────────────────────────────────────────────

def _section_macro(doc, d):
    _heading(doc, "ANALYSE MACRO — CONTEXTE INTERNATIONAL, AFRICAIN & UEMOA")

    for label, key in [
        ("Contexte international", "macro_international"),
        ("Contexte africain", "macro_africain"),
        ("Contexte UEMOA", "macro_uemoa"),
    ]:
        _heading(doc, label, 2)
        _para(doc, _no_data(_s(d, key)).replace("★", "").strip())

    _heading(doc, "Impacts sur le marché BRVM", 2)
    for label, key in [
        ("Impact sur les bourses mondiales", "impact_bourses_mondiales"),
        ("Impact sur les indicateurs BRVM", "impact_indicateurs_brvm"),
        ("Impact sur les sociétés cotées", "impact_societes_cotees"),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _bold(p, f"{label} : ", 10)
        _normal(p, _no_data(_s(d, key)).replace("★", "").strip(), 10)

    for label, key in [("Synthèse macro", "synthese_macro"), ("Recommandation", "recommandation_macro")]:
        val = _no_data(_s(d, key))
        if val != "Aucune information significative n'a été collectée sur ce point.":
            _heading(doc, label, 2)
            _para(doc, val.replace("★", "").strip())


# ── Section 5 — Actualités du marché BRVM ────────────────────────────────────

def _section_actualites(doc, d):
    _heading(doc, "ACTUALITÉS DU MARCHÉ BRVM")

    type_labels = {
        "AG": "Assemblées Générales",
        "AVIS": "Avis",
        "COMMUNIQUÉ": "Communiqués",
        "COMMUNIQUE": "Communiqués",
        "DIVIDENDE": "Dividendes",
        "DIVIDENDES": "Dividendes",
        "RÉSULTAT": "Résultats",
        "RESULTAT": "Résultats",
        "RÉSULTATS": "Résultats",
        "RESULTATS": "Résultats",
    }

    docs_off = _sl(d, "documents_officiels")
    if docs_off:
        grouped = defaultdict(list)
        for item in docs_off:
            grouped[str(item.get("type") or "Autre").upper()].append(item)

        for type_key, items in grouped.items():
            _heading(doc, type_labels.get(type_key, type_key.capitalize()), 2)
            for item in items:
                societe = item.get("societe") or "—"
                date_doc = item.get("date") or "—"
                detail = item.get("detail") or "—"
                _para(doc, f"Le {date_doc}, {societe} : {detail}")
    else:
        _para(doc, "Aucun document officiel recensé pour cette période.")

    alertes = _sl(d, "alertes_google")
    if alertes:
        _heading(doc, "Alertes Google", 2)
        _para(doc, "Les alertes suivantes ont été collectées et regroupées par thème.")

        theme_groups = defaultdict(list)
        for item in alertes:
            theme_groups[str(item.get("mot_cle") or "Général")].append(item)

        for theme, items in theme_groups.items():
            _heading(doc, theme, 3)
            for item in items:
                sent = str(item.get("sentiment", "neutre")).lower()
                emoji = "🟢" if sent == "positif" else ("🔴" if sent in ("negatif", "négatif") else "⚪")
                resume = item.get("resume") or "—"
                score = item.get("score_pertinence", "")
                texte = f"{emoji} {resume}"
                if score:
                    texte += f" (pertinence : {score}/10)"
                _para(doc, texte)
    else:
        _para(doc, "Aucune alerte Google recensée pour cette période.")


# ── Section 6 — Alertes du jour ──────────────────────────────────────────────

def _section_alertes(doc, d):
    _heading(doc, "ALERTES DU JOUR")

    _para(doc,
        "Les sociétés suivantes présentent des signaux extrêmes ou des incohérences "
        "majeures nécessitant une attention particulière."
    )

    surachat = _sl(d, "rsi_surachat")
    survente = _sl(d, "rsi_survente")
    divergences = _sl(d, "divergences_tech_fond")

    if surachat:
        _heading(doc, "Titres en surachat (RSI > 70)", 2)
        for item in surachat:
            sym = item.get("symbole", "—")
            rsi = item.get("rsi", "—")
            cours_max = item.get("cours_100j_max", "")
            signal_tech = item.get("signal_tech", "")
            signal_fond = item.get("signal_fond", "")
            reco = item.get("reco", "")
            detail = item.get("detail", "")

            texte = f"{sym} affiche un RSI en zone de surachat à {rsi}"
            if cours_max:
                texte += f", avec un cours au plus haut sur 100 jours à {cours_max} FCFA"
            if signal_tech and signal_fond:
                texte += (
                    f" et une divergence entre signal technique de {signal_tech} "
                    f"et fondamental de {signal_fond}"
                )
            elif detail:
                texte += f" — {detail}"
            if reco:
                texte += f" — recommandation maintenue à {reco}"
            _para(doc, texte + ".")

    if survente:
        _heading(doc, "Titres en survente (RSI < 30) — Opportunités potentielles", 2)
        for item in survente:
            sym = item.get("symbole", "—")
            rsi = item.get("rsi", "—")
            cours_min = item.get("cours_100j_min", "")
            signal_tech = item.get("signal_tech", "")
            signal_fond = item.get("signal_fond", "")
            reco = item.get("reco", "")
            detail = item.get("detail", "")

            texte = f"{sym} présente un RSI en zone de survente à {rsi}"
            if cours_min:
                texte += f", avec un cours proche de son plus bas sur 100 jours à {cours_min} FCFA"
            if signal_tech and signal_fond:
                texte += f", signal technique de {signal_tech} et fondamental de {signal_fond}"
            elif detail:
                texte += f" — {detail}"
            if reco:
                texte += f" — recommandation : {reco}"
            _para(doc, texte + ".")

    if divergences:
        _heading(doc, "Autres divergences techniques/fondamentales", 2)
        for item in divergences:
            sym = item.get("symbole", "—")
            reco = item.get("reco", "—")
            detail = item.get("detail", "—")
            _para(doc, f"{sym} [{reco}] : {detail}.")

    if not surachat and not survente and not divergences:
        _para(doc, "Aucune alerte majeure identifiée pour cette séance.")


# ── Section 7 — Conclusion ────────────────────────────────────────────────────

def _section_conclusion(doc, d, date_str: str):
    _heading(doc, "CONCLUSION")

    sentiment = _s(d, "conclusion_sentiment") or _s(d, "sentiment")
    _para(doc, f"Le marché BRVM affiche un sentiment global {sentiment}.")

    top3_opp = _sl(d, "top3_opportunites_conclusion") or _sl(d, "opportunites_majeures")
    if top3_opp:
        _heading(doc, "Top 3 opportunités du jour", 2)
        for i, item in enumerate(top3_opp[:3], 1):
            if isinstance(item, dict):
                sym = item.get("symbole", "—")
                raison = item.get("raison", "—")
                _para(doc, f"{i}. {sym} : {raison}")
            else:
                _para(doc, f"{i}. {item}")

    top3_risques = _sl(d, "top3_risques_conclusion") or _sl(d, "risques")
    if top3_risques:
        _heading(doc, "Top 3 risques à surveiller", 2)
        for i, item in enumerate(top3_risques[:3], 1):
            if isinstance(item, dict):
                risque = item.get("risque", "—")
                detail = item.get("detail", "")
                _para(doc, f"{i}. {risque}" + (f" : {detail}" if detail else ""))
            else:
                _para(doc, f"{i}. {item}")

    reco_finale = _no_data(_s(d, "recommandation_finale"))
    if reco_finale != "Aucune information significative n'a été collectée sur ce point.":
        _heading(doc, "Recommandation générale", 2)
        _para(doc, reco_finale)

    doc.add_paragraph()
    p_disc = doc.add_paragraph(
        "Cette note est produite à titre indicatif. Les analyses ne constituent pas des "
        "conseils en investissement. Tout investissement comporte des risques de perte en "
        "capital. Usage strictement réservé aux professionnels habilités."
    )
    for run in p_disc.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = _rgb("666666")
        run.italic = True

    doc.add_paragraph()
    p_sig = doc.add_paragraph()
    _bold(p_sig, f"Note générée le {date_str}  —  Agent GitHub Rapports BRVM", 9)
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER


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

    _section_entete_synthese(doc, data, date_str)
    _section_secteurs(doc, data)
    _section_liquidite(doc, data)
    _section_macro(doc, data)
    _section_actualites(doc, data)
    _section_alertes(doc, data)
    _section_conclusion(doc, data, date_str)
    _section_classement(doc, data)
    _section_portefeuilles(doc, data)

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
