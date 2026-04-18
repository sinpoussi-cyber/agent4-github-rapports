from datetime import date


COULEUR_HEADER = {
    "quotidien": "#1a73e8",
    "hebdo":     "#0f9d58",
    "mensuel":   "#e37400",
    "annuel":    "#7b1fa2",
}


def _score_color(score):
    if score <= 3:
        return "#0f9d58"
    if score <= 6:
        return "#e37400"
    return "#d93025"


def _bar_html(score):
    color = _score_color(score)
    pct = score * 10
    return (
        f'<div style="background:#e0e0e0;border-radius:4px;height:14px;width:200px;display:inline-block;">'
        f'<div style="background:{color};width:{pct}%;height:14px;border-radius:4px;"></div>'
        f'</div> <strong style="color:{color};">{score}/10</strong>'
    )


def generate(doc1_nom, doc2_nom, diff_data, analysis, type_rapport="quotidien"):
    """
    Génère le rapport email (subject, body_html, body_text) à partir des
    données de comparaison et de l'analyse Claude.
    """
    today = date.today()
    date_fr = today.strftime("%-d %B %Y") if hasattr(date, "strftime") else today.isoformat()
    # strftime "%-d" non dispo sur Windows — fallback propre
    try:
        date_fr = today.strftime("%-d %B %Y")
    except ValueError:
        date_fr = today.strftime("%d %B %Y").lstrip("0")

    label = {
        "quotidien": "quotidien",
        "hebdo": "hebdomadaire",
        "mensuel": "mensuel",
        "annuel": "annuel",
    }.get(type_rapport, type_rapport)

    annee = analysis.get("annee", today.year) if type_rapport == "annuel" else None
    if type_rapport == "annuel":
        subject = f"Rapport annuel comparaison BRVM — {annee}"
    else:
        subject = f"Rapport {label} — Comparaison rapports Word — {date_fr}"

    nb_ajoutes   = len(diff_data.get("paragraphes_ajoutes", []))
    nb_supprimes = len(diff_data.get("paragraphes_supprimes", []))
    nb_modifies  = len(diff_data.get("paragraphes_modifies", []))
    taux         = diff_data.get("taux_changement", 0)
    aucun_changement = (nb_ajoutes + nb_supprimes + nb_modifies) == 0

    header_color = COULEUR_HEADER.get(type_rapport, "#1a73e8")
    alerte       = str(analysis.get("alerte", "non")).lower() == "oui"
    score        = int(analysis.get("score_importance", 5))

    # ── HTML ────────────────────────────────────────────────────────────────
    alerte_badge = (
        '<div style="background:#d93025;color:#fff;padding:10px 18px;border-radius:6px;'
        'font-weight:bold;font-size:15px;margin-bottom:16px;">&#9888; ALERTE — Ces changements méritent une attention urgente</div>'
        if alerte else ""
    )

    if aucun_changement:
        contenu_html = '<p style="font-size:16px;color:#555;">✅ Aucun changement détecté entre les deux documents.</p>'
        contenu_text = "Aucun changement détecté entre les deux documents."
        evolution_html = ""
        evolution_text = ""
    else:
        changements_li = "".join(
            f'<li style="margin-bottom:6px;">{c}</li>'
            for c in analysis.get("changements_importants", [])
        )
        changements_text = "\n".join(
            f"  • {c}" for c in analysis.get("changements_importants", [])
        )

        evolution_html = ""
        evolution_text = ""
        if type_rapport == "annuel":
            nb_runs = analysis.get("nb_runs_annee", "—")
            date_debut = analysis.get("date_premier_run", "—")
            date_fin = analysis.get("date_dernier_run", "—")
            evolution_html = f"""
        <h2 style="color:#7b1fa2;margin-top:0;border-bottom:2px solid #ce93d8;padding-bottom:6px;">Evolution sur l'année {annee}</h2>
        <table style="border-collapse:collapse;width:100%;max-width:480px;margin-bottom:24px;">
          <tr style="background:#f3e5f5;">
            <td style="padding:8px 12px;border:1px solid #ce93d8;">Nombre de rapports analysés</td>
            <td style="padding:8px 12px;border:1px solid #ce93d8;font-weight:bold;color:#7b1fa2;">{nb_runs}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #ce93d8;">Premier rapport</td>
            <td style="padding:8px 12px;border:1px solid #ce93d8;">{date_debut}</td>
          </tr>
          <tr style="background:#f3e5f5;">
            <td style="padding:8px 12px;border:1px solid #ce93d8;">Dernier rapport</td>
            <td style="padding:8px 12px;border:1px solid #ce93d8;">{date_fin}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #ce93d8;">Taux de changement global</td>
            <td style="padding:8px 12px;border:1px solid #ce93d8;font-weight:bold;">{taux}%</td>
          </tr>
        </table>"""
            evolution_text = f"""EVOLUTION SUR L'ANNÉE {annee}
  Rapports analysés     : {nb_runs}
  Premier rapport       : {date_debut}
  Dernier rapport       : {date_fin}
  Taux changement global: {taux}%

"""

        contenu_html = f"""
        {evolution_html}
        {alerte_badge}

        <h2 style="color:#333;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">Statistiques des différences</h2>
        <table style="border-collapse:collapse;width:100%;max-width:480px;">
          <tr style="background:#f5f5f5;">
            <td style="padding:8px 12px;border:1px solid #ddd;">Paragraphes ajoutés</td>
            <td style="padding:8px 12px;border:1px solid #ddd;color:#0f9d58;font-weight:bold;">{nb_ajoutes}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #ddd;">Paragraphes supprimés</td>
            <td style="padding:8px 12px;border:1px solid #ddd;color:#d93025;font-weight:bold;">{nb_supprimes}</td>
          </tr>
          <tr style="background:#f5f5f5;">
            <td style="padding:8px 12px;border:1px solid #ddd;">Paragraphes modifiés</td>
            <td style="padding:8px 12px;border:1px solid #ddd;color:#e37400;font-weight:bold;">{nb_modifies}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #ddd;">Taux de changement</td>
            <td style="padding:8px 12px;border:1px solid #ddd;font-weight:bold;">{taux}%</td>
          </tr>
        </table>

        <h2 style="color:#333;margin-top:24px;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">Score d'importance</h2>
        <p>{_bar_html(score)}</p>

        <h2 style="color:#333;margin-top:24px;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">Résumé exécutif</h2>
        <p style="color:#444;line-height:1.6;">{analysis.get("resume_executif", "")}</p>

        <h2 style="color:#333;margin-top:24px;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">Changements importants</h2>
        <ul style="color:#444;line-height:1.8;">{changements_li}</ul>

        <h2 style="color:#333;margin-top:24px;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">Interprétation business</h2>
        <p style="color:#444;line-height:1.6;">{analysis.get("interpretation", "")}</p>
        """

        contenu_text = f"""{evolution_text}{"⚠ ALERTE — Ces changements méritent une attention urgente" + chr(10) if alerte else ""}
STATISTIQUES
  Paragraphes ajoutés   : {nb_ajoutes}
  Paragraphes supprimés : {nb_supprimes}
  Paragraphes modifiés  : {nb_modifies}
  Taux de changement    : {taux}%

SCORE D'IMPORTANCE : {score}/10

RÉSUMÉ EXÉCUTIF
{analysis.get("resume_executif", "")}

CHANGEMENTS IMPORTANTS
{changements_text}

INTERPRÉTATION BUSINESS
{analysis.get("interpretation", "")}
"""

    body_html = f"""<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f2f5;font-family:Arial,sans-serif;">
  <div style="max-width:680px;margin:32px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1);">

    <div style="background:{header_color};padding:28px 32px;">
      <h1 style="margin:0;color:#fff;font-size:22px;">Rapport {label.capitalize()}</h1>
      <p style="margin:6px 0 0;color:rgba(255,255,255,.85);font-size:14px;">{date_fr}</p>
    </div>

    <div style="padding:28px 32px;border-bottom:1px solid #e0e0e0;">
      <h2 style="margin:0 0 10px;color:#333;font-size:16px;">Documents comparés</h2>
      <p style="margin:4px 0;color:#555;">📄 <strong>Document 1 :</strong> {doc1_nom}</p>
      <p style="margin:4px 0;color:#555;">📄 <strong>Document 2 :</strong> {doc2_nom}</p>
    </div>

    <div style="padding:28px 32px;">
      {contenu_html}
    </div>

    <div style="background:#f5f5f5;padding:16px 32px;text-align:center;font-size:12px;color:#999;">
      Rapport généré automatiquement — Agent GitHub Rapports
    </div>

  </div>
</body>
</html>"""

    body_text = f"""RAPPORT {label.upper()} — {date_fr}
{"=" * 50}
Documents comparés :
  Document 1 : {doc1_nom}
  Document 2 : {doc2_nom}
{"=" * 50}
{contenu_text}
---
Rapport généré automatiquement — Agent GitHub Rapports
"""

    return {
        "subject":    subject,
        "body_html":  body_html,
        "body_text":  body_text,
    }
