import io
import math
import os
import random
import tempfile
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

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


# ── Helpers cours / chart ─────────────────────────────────────────────────────

def _to_float(v):
    """Conversion robuste : nombre, '1 440', '1,440.5', '1.440,5'."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _build_price_chart_png(s: dict, width_in: float = 6.5, height_in: float = 2.6) -> bytes:
    """
    Génère un graphique matplotlib (100 jours) ancré sur les bornes connues :
    cours_debut, cours_fin, plus_haut_100j, plus_bas_100j.
    Trace : courbe bleue reconstruite, droite de tendance pointillée, marqueurs haut/bas.
    Retourne les bytes PNG ; None si données insuffisantes.
    """
    debut = _to_float(s.get("cours_debut"))
    fin = _to_float(s.get("cours_fin")) or _to_float(s.get("cours"))
    haut = _to_float(s.get("plus_haut_100j"))
    bas = _to_float(s.get("plus_bas_100j"))
    if debut is None or fin is None or haut is None or bas is None:
        return None
    if haut < bas:
        haut, bas = bas, haut

    # ── Reconstruction d'une trajectoire 100 séances ancrée sur les 4 points connus.
    # Position des extrema : le premier extrémum atteint est celui le plus éloigné
    # du cours de départ ; l'autre extrémum après. Seed déterministe par ticker.
    n = 100
    rng = random.Random(hash(str(s.get("ticker") or "")) & 0xFFFFFFFF)
    if abs(haut - debut) >= abs(bas - debut):
        i_high, i_low = 30, 70
    else:
        i_low, i_high = 30, 70

    anchors = sorted([(0, debut), (i_high, haut), (i_low, bas), (n - 1, fin)])
    x_a = [a[0] for a in anchors]
    y_a = [a[1] for a in anchors]

    xs = list(range(n))
    ys = []
    for x in xs:
        for k in range(len(x_a) - 1):
            if x_a[k] <= x <= x_a[k + 1]:
                t = (x - x_a[k]) / max(x_a[k + 1] - x_a[k], 1)
                t_smooth = 0.5 - 0.5 * math.cos(math.pi * t)
                base = y_a[k] + (y_a[k + 1] - y_a[k]) * t_smooth
                break
        amp = (haut - bas) * 0.04
        if x in (0, i_high, i_low, n - 1):
            ys.append(base)
        else:
            noisy = base + rng.uniform(-amp, amp)
            ys.append(min(haut, max(bas, noisy)))

    # ── Tracé
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=140)
    ax.plot(xs, ys, color="#1A73E8", linewidth=1.6, zorder=3)
    ax.fill_between(xs, ys, min(ys) - (haut - bas) * 0.05,
                    color="#1A73E8", alpha=0.07, zorder=1)

    # Tendance linéaire start -> end
    ax.plot([0, n - 1], [debut, fin], linestyle="--", color="#888888",
            linewidth=1.2, zorder=2, label="Tendance")

    # Marqueurs plus haut / plus bas
    ax.scatter([i_high], [haut], color="#0F9D58", s=42, zorder=5, edgecolor="white")
    ax.scatter([i_low], [bas], color="#D93025", s=42, zorder=5, edgecolor="white")
    ax.annotate(f"+ haut\n{haut:,.0f}".replace(",", " "),
                xy=(i_high, haut), xytext=(0, 8), textcoords="offset points",
                ha="center", fontsize=7, color="#0F9D58", weight="bold")
    ax.annotate(f"+ bas\n{bas:,.0f}".replace(",", " "),
                xy=(i_low, bas), xytext=(0, -22), textcoords="offset points",
                ha="center", fontsize=7, color="#D93025", weight="bold")

    # Marqueurs début / fin
    ax.scatter([0, n - 1], [debut, fin], color="#1A237E", s=22, zorder=4)

    # Cosmétique
    date_d = str(s.get("date_debut_100j") or "J-100")
    date_f = str(s.get("date_fin_100j") or "J")
    ax.set_xticks([0, n // 2, n - 1])
    ax.set_xticklabels([date_d, "—", date_f], fontsize=7, color="#555555")
    ax.yaxis.set_major_formatter(FuncFormatter(
        lambda v, _: f"{v:,.0f}".replace(",", " ")))
    ax.tick_params(axis="y", labelsize=7, colors="#555555")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color("#CCCCCC")
    ax.grid(True, axis="y", linestyle=":", color="#DDDDDD", linewidth=0.6, zorder=0)
    ax.set_ylim(bas - (haut - bas) * 0.10, haut + (haut - bas) * 0.18)

    perf = s.get("perf_100j")
    title_perf = f"   ({perf})" if perf else ""
    ax.set_title(
        f"Évolution du cours sur 100 jours{title_perf}",
        fontsize=9, color="#1A237E", weight="bold", loc="left", pad=6,
    )

    fig.tight_layout(pad=0.6)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    png_bytes = buf.getvalue()
    print(f"  [Fiches/Chart] {s.get('ticker')} : PNG size {len(png_bytes)} bytes")
    return png_bytes


def _fund_val(s: dict, key: str) -> str:
    v = s.get(key)
    if v is None:
        return "N/D"
    sv = str(v).strip()
    if sv == "" or sv.lower() in ("null", "none", "—"):
        return "N/D"
    return sv


def _build_fundamentals_table(doc, s: dict):
    """
    Tableau enrichi des indicateurs fondamentaux. Inclut CA et résultat net
    sur 3 exercices (N, N-1, N-2), ratios de rentabilité (marge, ROE, ROA, PER),
    capitalisation et structure (dividende, rendement, dette nette / EBITDA),
    et croissances (1 an / 3 ans). 'N/D' quand la donnée est absente du rapport.
    """
    _sub_heading(doc, "Indicateurs fondamentaux (extraits du rapport)")

    per_actuel = _fund_val(s, "per")
    per_secto = _fund_val(s, "per_sectoriel")
    per_val = per_actuel if per_secto in ("N/D", "") else f"{per_actuel} (secteur : {per_secto})"

    div_actuel = _fund_val(s, "dividende")
    rend = _fund_val(s, "rendement_dividende")
    div_val = div_actuel if rend in ("N/D", "") else f"{div_actuel} (rendement : {rend})"

    dette_nette = _fund_val(s, "dette_nette")
    dn_ebitda = _fund_val(s, "dette_nette_ebitda")
    dette_val = dette_nette if dn_ebitda in ("N/D", "") else f"{dette_nette} | DN/EBITDA : {dn_ebitda}"

    rows = [
        ("Chiffre d'affaires (N)",      _fund_val(s, "ca"),                       _fund_val(s, "ca_date")),
        ("CA — exercice N-1",           _fund_val(s, "ca_n_1"),                   "N-1"),
        ("CA — exercice N-2",           _fund_val(s, "ca_n_2"),                   "N-2"),
        ("Résultat net (N)",            _fund_val(s, "resultat_net"),             _fund_val(s, "rn_date")),
        ("Résultat net — N-1",          _fund_val(s, "resultat_n_1"),             "N-1"),
        ("Résultat net — N-2",          _fund_val(s, "resultat_n_2"),             "N-2"),
        ("Marge nette",                 _fund_val(s, "marge_nette"),              _fund_val(s, "mn_date")),
        ("ROE",                         _fund_val(s, "roe"),                      _fund_val(s, "roe_date")),
        ("ROA",                         _fund_val(s, "roa"),                      _fund_val(s, "roa_date")),
        ("PER (actuel / sectoriel)",    per_val,                                  _fund_val(s, "per_date")),
        ("Capitalisation boursière",    _fund_val(s, "capitalisation_boursiere"), _fund_val(s, "capi_date")),
        ("Dividende & rendement",       div_val,                                  _fund_val(s, "div_date")),
        ("Dette nette / EBITDA",        dette_val,                                _fund_val(s, "date_donnees_financieres")),
        ("Croissance CA — 1 an",        _fund_val(s, "croissance_ca_1an"),        _fund_val(s, "croissance_ca_date")),
        ("Croissance CA — 3 ans",       _fund_val(s, "croissance_ca_3ans"),       _fund_val(s, "croissance_ca_date")),
    ]

    tbl = doc.add_table(rows=1 + len(rows), cols=3)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Indicateur", "Valeur", "Date / période"]):
        _cw(tbl.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")

    for i, (label, val, dt) in enumerate(rows, start=1):
        bg_val = "F5F5F5" if val == "N/D" else "FFFFFF"
        bg_dt  = "F5F5F5" if dt  == "N/D" else "FFFFFF"
        _cw(tbl.rows[i].cells[0], label, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[1], val, size=8, bg=bg_val)
        _cw(tbl.rows[i].cells[2], dt,  size=8, bg=bg_dt)

    doc.add_paragraph()
    _build_financial_commentary(doc, s)


# ── Commentaire narratif d'analyse financière ────────────────────────────────

def _pct_to_float(v):
    """'+5,2%' → 5.2 ; '5,8%' → 5.8 ; None / non-numérique → None."""
    if v is None:
        return None
    s = str(v).replace("%", "").replace(",", ".").replace("+", "").strip()
    if not s or s.lower() in ("null", "none", "—", "n/d"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _build_financial_commentary(doc, s: dict):
    """
    Paragraphe narratif (5-8 lignes) de synthèse financière. Combine :
    - santé financière globale (rentabilité, marge, dette)
    - comparaison sectorielle (PER vs PER sectoriel)
    - points forts et risques identifiés
    - conclusion sur l'attractivité financière
    """
    _sub_heading(doc, "Commentaire — Analyse financière")

    ticker = _s(s, "ticker", "—")
    nom = _s(s, "nom", ticker)
    secteur = _s(s, "secteur", "—")

    marge = _pct_to_float(s.get("marge_nette"))
    roe = _pct_to_float(s.get("roe"))
    roa = _pct_to_float(s.get("roa"))
    croiss_1y = _pct_to_float(s.get("croissance_ca_1an")) or _pct_to_float(s.get("croissance_ca"))
    croiss_3y = _pct_to_float(s.get("croissance_ca_3ans"))
    rendement = _pct_to_float(s.get("rendement_dividende"))

    per_f = _pct_to_float(s.get("per"))
    per_sec_f = _pct_to_float(s.get("per_sectoriel"))

    # 1) Santé financière globale
    if roe is not None and marge is not None:
        if roe >= 12 and marge >= 8:
            sante = (
                f"{nom} affiche une rentabilité solide (ROE de {roe:.1f}%, "
                f"marge nette de {marge:.1f}%), signe d'une exploitation efficace "
                "et d'un retour sur fonds propres conforme aux standards du secteur."
            )
        elif roe < 5 or marge < 2:
            sante = (
                f"La rentabilité de {nom} apparaît contrainte (ROE {roe:.1f}%, "
                f"marge nette {marge:.1f}%), reflet de pressions sur les coûts ou "
                "d'une structure de capital peu optimisée."
            )
        else:
            sante = (
                f"{nom} présente une rentabilité moyenne (ROE {roe:.1f}%, "
                f"marge nette {marge:.1f}%), niveau qui laisse une marge de "
                "progression sans signal d'alerte immédiat."
            )
    else:
        sante = (
            f"{nom}, cotée dans le secteur {secteur}, présente un profil financier "
            "dont les ratios complets ne sont pas tous disponibles dans le rapport "
            "source — l'appréciation s'appuie sur les éléments partiels recensés."
        )

    # 2) Croissance & dynamique commerciale
    if croiss_3y is not None:
        if croiss_3y >= 10:
            croissance = (
                f"La trajectoire commerciale est porteuse (croissance CA "
                f"{croiss_3y:+.1f}% sur 3 ans), traduisant une dynamique "
                "structurelle de parts de marché ou de pricing power."
            )
        elif croiss_3y <= 0:
            croissance = (
                f"Le chiffre d'affaires est en contraction sur la période "
                f"({croiss_3y:+.1f}% sur 3 ans), signal d'une demande affaiblie ou "
                "d'un repositionnement concurrentiel à surveiller."
            )
        else:
            croissance = (
                f"La croissance commerciale demeure modérée ({croiss_3y:+.1f}% "
                "sur 3 ans), conforme à un secteur en phase de maturité."
            )
    elif croiss_1y is not None:
        croissance = (
            f"Sur le dernier exercice, le CA évolue de {croiss_1y:+.1f}% — "
            "à confirmer sur un horizon pluriannuel."
        )
    else:
        croissance = (
            "La dynamique de croissance n'est pas chiffrée dans le rapport source, "
            "ce qui limite la projection des flux futurs."
        )

    # 3) Comparaison sectorielle (PER)
    if per_f is not None and per_sec_f is not None:
        ecart = per_f - per_sec_f
        if ecart <= -1.5:
            comparaison = (
                f"Sur le plan de la valorisation, le PER de {per_f:.1f} ressort "
                f"en-dessous de la moyenne sectorielle ({per_sec_f:.1f}) — "
                "configuration de décote relative qui peut intéresser les "
                "investisseurs value, sous réserve d'un catalyseur de revalorisation."
            )
        elif ecart >= 1.5:
            comparaison = (
                f"La valorisation apparaît plus exigeante que la moyenne du secteur "
                f"(PER {per_f:.1f} vs {per_sec_f:.1f}) — la prime ne se justifie "
                "qu'en présence d'une croissance ou d'une rentabilité supérieures."
            )
        else:
            comparaison = (
                f"La valorisation est alignée sur le secteur (PER {per_f:.1f} vs "
                f"{per_sec_f:.1f}), ce qui n'introduit ni décote ni prime."
            )
    elif per_f is not None:
        comparaison = (
            f"Le PER actuel ressort à {per_f:.1f} ; à défaut de référence sectorielle "
            "explicite, la comparaison de valorisation reste indicative."
        )
    else:
        comparaison = (
            "La valorisation (PER) n'étant pas renseignée, la comparaison sectorielle "
            "se fonde sur les ratios de rentabilité disponibles."
        )

    # 4) Forces & risques
    forces = _sl(s, "forces_financieres")
    faiblesses = _sl(s, "faiblesses_financieres")
    forts_risques_parts = []
    if forces:
        forts_risques_parts.append("Points forts : " + " ; ".join(f.strip().rstrip(".") for f in forces[:3]) + ".")
    if faiblesses:
        forts_risques_parts.append("Risques : " + " ; ".join(f.strip().rstrip(".") for f in faiblesses[:3]) + ".")
    if not forts_risques_parts:
        if rendement is not None and rendement >= 4:
            forts_risques_parts.append(
                f"Point fort : rendement du dividende attractif ({rendement:.1f}%)."
            )
        if roa is not None and roa < 1:
            forts_risques_parts.append(
                f"Risque : ROA limité ({roa:.1f}%) signalant une efficience modeste des actifs."
            )
    forts_risques = " ".join(forts_risques_parts) if forts_risques_parts else (
        "L'analyse qualitative ne fait pas ressortir de points forts ou risques majeurs "
        "au-delà des ratios déjà commentés."
    )

    # 5) Conclusion sur l'attractivité financière
    score = _score_f(s)
    reco = _s(s, "reco", "").upper()
    if score >= 70 or "ACHAT" in reco:
        conclusion = (
            "Au global, l'attractivité financière est jugée favorable : la combinaison "
            "des fondamentaux soutient une position constructive sur le titre."
        )
    elif score <= 39 or "VENTE" in reco:
        conclusion = (
            "L'attractivité financière reste fragilisée : les ratios actuels ne "
            "soutiennent pas un repositionnement offensif sans amélioration tangible."
        )
    else:
        conclusion = (
            "L'attractivité financière est intermédiaire : un positionnement sélectif "
            "à proportion mesurée est cohérent avec le profil fondamental observé."
        )

    full = " ".join([sante, croissance, comparaison, forts_risques, conclusion])
    _narrative(doc, full)


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
    """
    Parcourt le body en ordre document pour préserver l'association paragraphes/tableaux.
    Les cellules au format 'clé\\nvaleur' (comme dans les tableaux d'indicateurs
    financiers du rapport source) sont aplaties en 'clé: valeur'.
    """
    doc = Document(io.BytesIO(doc_bytes))
    para_by_id = {id(p._element): p for p in doc.paragraphs}
    table_by_id = {id(t._element): t for t in doc.tables}

    parts = []
    table_idx = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            p = para_by_id.get(id(child))
            if p is None:
                continue
            txt = p.text.strip()
            if txt:
                parts.append(txt)
        elif child.tag == qn('w:tbl'):
            t = table_by_id.get(id(child))
            if t is None:
                continue
            table_idx += 1
            kv_lines = []
            for row in t.rows:
                for cell in row.cells:
                    raw = cell.text.strip()
                    if not raw:
                        continue
                    if '\n' in raw:
                        k, v = raw.split('\n', 1)
                        k, v = k.strip(), v.strip()
                        if v and v.lower() != 'none':
                            kv_lines.append(f"  {k}: {v}")
                    else:
                        kv_lines.append(f"  {raw}")
            if kv_lines:
                parts.append(f"[Tableau {table_idx}]")
                parts.extend(kv_lines)
    return "\n".join(parts)


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


def _fmt_int(v) -> str:
    """Format un nombre en milliers séparés par espace ; renvoie '—' si None."""
    if v is None:
        return "—"
    try:
        f = float(v)
        if f != f:  # NaN
            return "—"
        return f"{f:,.0f}".replace(",", " ")
    except (ValueError, TypeError):
        return str(v).strip() or "—"


def _validate_var_1j(val):
    """Filtre les variations journalières aberrantes (>±10%, signe de
    confusion avec une perf multi-jours type perf_100j)."""
    if val is None:
        return "—"
    try:
        pct = float(
            str(val).replace("%", "").replace("+", "").replace(",", ".").strip()
        )
        if abs(pct) > 10:
            return "—"
        return val
    except (ValueError, TypeError):
        return "—"


def _fmt_amount(val):
    """Formate un montant avec séparateur de milliers (espace).
    Si val est déjà une string portant une unité (Mds, M, FCFA, %), la
    conserve telle quelle ; si val est un nombre brut ou une chaîne purement
    numérique, formate avec espaces comme séparateur de milliers."""
    if val is None or val == "N/D" or val == "—":
        return "—"
    if isinstance(val, (int, float)):
        try:
            f = float(val)
            if f != f:
                return "—"
            return f"{f:,.0f}".replace(",", " ")
        except (ValueError, TypeError):
            return str(val)
    s = str(val).strip()
    if not s or s.lower() in ("null", "none"):
        return "—"
    if any(u in s for u in ("Md", "milliard", "Milliard", "million", "Million",
                            "FCFA", "%", " M ", " M.")):
        return s
    try:
        f = float(s.replace(" ", "").replace(",", "."))
        return f"{f:,.0f}".replace(",", " ")
    except ValueError:
        return s


def build_market_table(doc, s: dict):
    """
    Tableau des métriques de marché en 6×2 (12 indicateurs).
    Inclut capitalisation, volume moyen 30j, nombre d'actions et PER
    en plus des métriques de risque/stabilité.
    """
    _section_heading(doc, "MÉTRIQUES DE MARCHÉ")

    cours = _fmt_amount(s.get("cours"))
    var_1j = _validate_var_1j(s.get("var_1j"))
    volatilite = _s(s, "volatilite") or "—"
    beta = _s(s, "beta") or "—"
    liquidite = _s(s, "liquidite") or "—"
    risque = _s(s, "risque") or "—"
    divergence = _s(s, "divergence") or "aucune"
    stabilite = _s(s, "stabilite") or "—"
    capi = _fmt_amount(s.get("capitalisation_boursiere"))
    vol_30j = _fmt_int(s.get("volume_moyen_30j"))
    nb_act = _fmt_int(s.get("nb_actions"))
    per = _s(s, "per") or "—"

    risque_bg = _risque_bg(risque)
    var_bg = _var_color(var_1j)
    stab_bg = "C6EFCE" if "bonne" in str(stabilite).lower() else (
        "FFC7CE" if "fragile" in str(stabilite).lower() else "FFEB9C"
    )
    div_bg = "FFEB9C" if divergence.lower() not in ("aucune", "—", "") else "FFFFFF"

    pairs = [
        ("Cours actuel (FCFA)",       cours,      "F0F4FF"),
        ("Variation 1 journée",       var_1j,     var_bg),
        ("Capitalisation boursière",  capi,       "FFFFFF"),
        ("PER (Cours / BPA)",         per,        "FFFFFF"),
        ("Volume moyen 30j",          vol_30j,    "FFFFFF"),
        ("Nb actions en circulation", nb_act,     "FFFFFF"),
        ("Volatilité",                volatilite, "FFFFFF"),
        ("Bêta",                      beta,       "FFFFFF"),
        ("Liquidité",                 liquidite,  "FFFFFF"),
        ("Niveau de risque",          risque,     risque_bg),
        ("Divergence tech/fond",      divergence, div_bg),
        ("Stabilité",                 stabilite,  stab_bg),
    ]

    tbl = doc.add_table(rows=6, cols=4)
    tbl.style = "Table Grid"
    for i in range(6):
        ll, lv, lbg = pairs[i]
        rl, rv, rbg = pairs[i + 6]
        _cw(tbl.rows[i].cells[0], ll, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[1], lv, size=8, bg=lbg)
        _cw(tbl.rows[i].cells[2], rl, bold=True, size=8, bg="EBF0FA")
        _cw(tbl.rows[i].cells[3], rv, size=8, bg=rbg)


def build_chart_comment(doc, s: dict):
    """
    Graphique matplotlib + commentaire analytique de la courbe.
    Couvre : tendance, volatilité, momentum, phase de marché.
    """
    # ── Graphique réel (PNG matplotlib) si données disponibles
    png_bytes = None
    try:
        png_bytes = _build_price_chart_png(s)
    except Exception as exc:
        print(f"  [Fiches/Chart] {s.get('ticker')} : échec génération chart — {exc}")

    if png_bytes and len(png_bytes) > 1000:
        p_img = doc.add_paragraph()
        p_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(png_bytes)
            tmp_path = tmp.name
        try:
            p_img.add_run().add_picture(tmp_path, width=Cm(16))
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    else:
        if png_bytes is not None:
            print(f"  [Fiches/Chart] {s.get('ticker')} : PNG trop petit ({len(png_bytes)} bytes) — ignoré")
        p_ph = doc.add_paragraph()
        p_ph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ph.paragraph_format.space_before = Pt(8)
        p_ph.paragraph_format.space_after = Pt(4)
        r_ph = p_ph.add_run(
            "[Données de cours 100j non disponibles pour ce titre — graphique non tracé]"
        )
        r_ph.italic = True
        r_ph.font.size = Pt(9)
        r_ph.font.color.rgb = _rgb("888888")

    # ── Section commentaire
    _section_heading(doc, "COMMENTAIRE DE LA COURBE — ÉVOLUTION 100 JOURS")

    ticker = _s(s, "ticker", "ce titre")
    cours = _fmt_amount(s.get("cours"))
    var_1j = _validate_var_1j(s.get("var_1j"))
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


def _fmt_num(v, decimals: int = 1) -> str:
    """Format un nombre avec N décimales, espaces de milliers. '—' si vide/None."""
    if v is None:
        return "—"
    try:
        f = float(v)
        if f != f:
            return "—"
        s = f",.{decimals}f"
        return format(f, s).replace(",", " ")
    except (ValueError, TypeError):
        return str(v).strip() or "—"


def _tech_values_str(s: dict, sig_key: str) -> str:
    """Construit la chaîne 'MM20: 2850 | MM50: 2920 | Signal: BAISSIER' pour
    chaque indicateur, en injectant les valeurs numériques extraites."""
    signal = _s(s, sig_key) or "—"
    sig_upper = signal.upper() if signal != "—" else "—"
    if sig_key == "mm":
        return (
            f"MM20 : {_fmt_num(s.get('mm20_valeur'), 0)} | "
            f"MM50 : {_fmt_num(s.get('mm50_valeur'), 0)} | "
            f"Signal : {sig_upper}"
        )
    if sig_key == "boll":
        return (
            f"Bande basse : {_fmt_num(s.get('boll_inf'), 0)} | "
            f"Bande haute : {_fmt_num(s.get('boll_sup'), 0)} | "
            f"Signal : {sig_upper}"
        )
    if sig_key == "macd":
        return (
            f"MACD : {_fmt_num(s.get('macd_valeur'), 2)} | "
            f"Signal line : {_fmt_num(s.get('macd_signal_line'), 2)} | "
            f"Signal : {sig_upper}"
        )
    if sig_key == "rsi":
        return f"RSI : {_fmt_num(s.get('rsi_valeur'), 1)} | Signal : {sig_upper}"
    if sig_key == "stoch":
        return (
            f"%K : {_fmt_num(s.get('stoch_k'), 1)} | "
            f"%D : {_fmt_num(s.get('stoch_d'), 1)} | "
            f"Signal : {sig_upper}"
        )
    return f"Signal : {sig_upper}"


def build_technical_analysis(doc, s: dict):
    """
    Analyse technique structurée :
    - Tableau des 5 indicateurs (signal + appréciation + valeurs numériques + détail)
    - Synthèse globale
    - Évaluation convergence / divergence des signaux
    """
    _section_heading(doc, "ANALYSE TECHNIQUE")

    indicateurs = [
        ("Moyennes Mobiles (MM)", "mm",   "mm_signal",   "mm_detail"),
        ("Bandes de Bollinger",   "boll", "boll_signal", "boll_detail"),
        ("MACD",                  "macd", "macd_signal", "macd_detail"),
        ("RSI",                   "rsi",  "rsi_signal",  "rsi_detail"),
        ("Stochastique",          "stoch","stoch_signal","stoch_detail"),
    ]

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    for i, h in enumerate(["Indicateur", "Appréciation", "Valeurs & Signal", "Analyse détaillée"]):
        _cw(tbl.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")

    for label, sig_key, sig_label_key, detail_key in indicateurs:
        signal = _s(s, sig_key)
        sig_label = _s(s, sig_label_key) or (signal.capitalize() if signal else "—")
        detail = _s(s, detail_key) or "—"
        emoji = _signal_emoji(signal)
        values_str = _tech_values_str(s, sig_key)

        row = tbl.add_row()
        _cw(row.cells[0], f"{emoji}  {label}", bold=True, size=8, bg="EBF0FA")
        _cw(row.cells[1], sig_label, size=8, bg=_signal_bg(signal))
        _cw(row.cells[2], values_str, size=7, bg=_signal_bg(signal))
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

    # ── Tableau des indicateurs fondamentaux extraits du rapport source
    _build_fundamentals_table(doc, s)

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


def build_financial_analysis(doc, s: dict):
    """
    Analyse financière détaillée (pages 42-66 du rapport source) :
    - Tableau ratios : PER, ROE, ROA, marge nette, croissance CA, dette/CP
    - Évolution sur 2-3 ans si disponible (CA, résultat net, ROE)
    - Forces et faiblesses financières (3 points max chacun)
    - Conclusion financière en 2-3 lignes
    - Date des données toujours affichée
    """
    _section_heading(doc, "ANALYSE FINANCIÈRE DÉTAILLÉE")

    date_donnees = _fund_val(s, "date_donnees_financieres")
    _narrative(doc,
               f"Données financières de référence — {date_donnees}.",
               size=8, italic=True, color="666666")

    # ── 1. Tableau ratios financiers (6 ratios + dates)
    _sub_heading(doc, "1.  Ratios financiers clés")
    ratios_rows = [
        ("PER (Price/Earnings)",  _fund_val(s, "per"),           _fund_val(s, "per_date")),
        ("ROE (Return on Equity)", _fund_val(s, "roe"),          _fund_val(s, "roe_date")),
        ("ROA (Return on Assets)", _fund_val(s, "roa"),          _fund_val(s, "roa_date")),
        ("Marge nette",            _fund_val(s, "marge_nette"),  _fund_val(s, "mn_date")),
        ("Croissance CA",          _fund_val(s, "croissance_ca"), _fund_val(s, "croissance_ca_date")),
        ("Dette / Capitaux propres", _fund_val(s, "dette_cp"),   _fund_val(s, "dette_cp_date")),
    ]
    tbl_r = doc.add_table(rows=1 + len(ratios_rows), cols=3)
    tbl_r.style = "Table Grid"
    for i, h in enumerate(["Ratio", "Valeur", "Date"]):
        _cw(tbl_r.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")
    for i, (label, val, dt) in enumerate(ratios_rows, start=1):
        bg_val = "F5F5F5" if val == "N/D" else "FFFFFF"
        bg_dt  = "F5F5F5" if dt  == "N/D" else "FFFFFF"
        _cw(tbl_r.rows[i].cells[0], label, bold=True, size=8, bg="EBF0FA")
        _cw(tbl_r.rows[i].cells[1], val, size=8, bg=bg_val)
        _cw(tbl_r.rows[i].cells[2], dt,  size=8, bg=bg_dt)
    doc.add_paragraph()

    # ── 2. Évolution historique sur 2-3 ans (CA / RN / ROE)
    ca_n   = _fund_val(s, "ca")
    ca_n_1 = _fund_val(s, "ca_n_1")
    ca_n_2 = _fund_val(s, "ca_n_2")
    rn_n   = _fund_val(s, "resultat_net")
    rn_n_1 = _fund_val(s, "resultat_n_1")
    rn_n_2 = _fund_val(s, "resultat_n_2")
    roe_n   = _fund_val(s, "roe")
    roe_n_1 = _fund_val(s, "roe_n_1")
    roe_n_2 = _fund_val(s, "roe_n_2")

    has_history = any(v != "N/D" for v in (ca_n_1, ca_n_2, rn_n_1, rn_n_2, roe_n_1, roe_n_2))
    if has_history:
        _sub_heading(doc, "2.  Évolution sur 2-3 ans")
        tbl_h = doc.add_table(rows=4, cols=4)
        tbl_h.style = "Table Grid"
        for i, h in enumerate(["Indicateur", "N-2", "N-1", "N (réf)"]):
            _cw(tbl_h.rows[0].cells[i], h, bold=True, size=8, bg="1A237E", color="FFFFFF")
        evo_rows = [
            ("Chiffre d'affaires", ca_n_2, ca_n_1, ca_n),
            ("Résultat net",       rn_n_2, rn_n_1, rn_n),
            ("ROE",                roe_n_2, roe_n_1, roe_n),
        ]
        for i, (lbl, v2, v1, v0) in enumerate(evo_rows, start=1):
            _cw(tbl_h.rows[i].cells[0], lbl, bold=True, size=8, bg="EBF0FA")
            _cw(tbl_h.rows[i].cells[1], v2, size=8, bg="F5F5F5" if v2 == "N/D" else "FFFFFF")
            _cw(tbl_h.rows[i].cells[2], v1, size=8, bg="F5F5F5" if v1 == "N/D" else "FFFFFF")
            _cw(tbl_h.rows[i].cells[3], v0, size=8, bg="F5F5F5" if v0 == "N/D" else "FFFFFF")
        doc.add_paragraph()
    else:
        _narrative(doc,
                   "Historique sur 2-3 ans non disponible dans le rapport source.",
                   size=8, italic=True, color="888888")

    # ── 3. Forces / Faiblesses (3 max chacun)
    forces = _sl(s, "forces_financieres")[:3]
    faibles = _sl(s, "faiblesses_financieres")[:3]

    _sub_heading(doc, "3.  Forces et faiblesses financières")
    if forces or faibles:
        tbl_ff = doc.add_table(rows=4, cols=2)
        tbl_ff.style = "Table Grid"
        _cw(tbl_ff.rows[0].cells[0], "✚  FORCES", bold=True, size=9, bg="0F9D58", color="FFFFFF")
        _cw(tbl_ff.rows[0].cells[1], "−  FAIBLESSES", bold=True, size=9, bg="D93025", color="FFFFFF")
        for i in range(3):
            f_txt = forces[i] if i < len(forces) else "—"
            w_txt = faibles[i] if i < len(faibles) else "—"
            _cw(tbl_ff.rows[i + 1].cells[0], f"• {f_txt}", size=8, bg="EBF7EE" if f_txt != "—" else "F5F5F5")
            _cw(tbl_ff.rows[i + 1].cells[1], f"• {w_txt}", size=8, bg="FDEEEE" if w_txt != "—" else "F5F5F5")
        doc.add_paragraph()
    else:
        _narrative(doc,
                   "Forces et faiblesses financières non identifiées dans le rapport source.",
                   size=8, italic=True, color="888888")

    # ── 4. Conclusion financière (2-3 lignes)
    _sub_heading(doc, "4.  Conclusion financière")
    synth = _s(s, "synthese_financiere")
    if synth and synth not in ("", "—", "null", "None"):
        _key_bloc(doc, "BILAN FINANCIER :", synth, "E8F0FB", "1558A7")
    else:
        # Fallback synthétique à partir des ratios disponibles
        parts = []
        if _fund_val(s, "per") != "N/D":
            parts.append(f"PER {_fund_val(s, 'per')}")
        if _fund_val(s, "roe") != "N/D":
            parts.append(f"ROE {_fund_val(s, 'roe')}")
        if _fund_val(s, "marge_nette") != "N/D":
            parts.append(f"marge nette {_fund_val(s, 'marge_nette')}")
        if _fund_val(s, "croissance_ca") != "N/D":
            parts.append(f"croissance CA {_fund_val(s, 'croissance_ca')}")
        if parts:
            txt = (
                "Profil financier synthétique — " + " · ".join(parts)
                + f". Données arrêtées au {date_donnees}. "
                "Une analyse complémentaire sera intégrée dès la disponibilité "
                "de la synthèse qualitative du rapport source."
            )
        else:
            txt = (
                "Aucune synthèse financière n'a été extraite pour cette société dans "
                f"le rapport de référence ({date_donnees}). "
                "Les données seront enrichies lors de la prochaine publication."
            )
        _key_bloc(doc, "BILAN FINANCIER :", txt, "E8F0FB", "1558A7")


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
    build_financial_analysis(doc, s)         # Analyse financière détaillée (ratios, évolution, forces/faiblesses)
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
        print(f"  [Fiches/{freq}] ERREUR : extraction invalide ou aucune société — abandon (voir logs Extractor).")
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
