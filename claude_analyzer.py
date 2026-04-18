import os
import json
import anthropic
from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")


def _build_prompt(doc1_nom, doc2_nom, diff_data):
    ajoutes = "\n".join(f"  - {p}" for p in diff_data.get("paragraphes_ajoutes", [])) or "  Aucun"
    supprimes = "\n".join(f"  - {p}" for p in diff_data.get("paragraphes_supprimes", [])) or "  Aucun"
    modifies = "\n".join(
        f"  - AVANT: {m['avant']}\n    APRÈS: {m['apres']} (similarité: {m['similarite']})"
        for m in diff_data.get("paragraphes_modifies", [])
    ) or "  Aucun"

    doc1_stats = diff_data.get("doc1_stats", {})
    doc2_stats = diff_data.get("doc2_stats", {})

    return f"""Tu es un analyste expert en rapports d'activité. Analyse les différences entre deux rapports Word.

DOCUMENT 1 : {doc1_nom}
  - Paragraphes : {doc1_stats.get('nb_paragraphes')} | Mots : {doc1_stats.get('nb_mots')} | Caractères : {doc1_stats.get('nb_caracteres')}

DOCUMENT 2 : {doc2_nom}
  - Paragraphes : {doc2_stats.get('nb_paragraphes')} | Mots : {doc2_stats.get('nb_mots')} | Caractères : {doc2_stats.get('nb_caracteres')}

RÉSUMÉ DES DIFFÉRENCES : {diff_data.get('resume_changements')}
TAUX DE CHANGEMENT : {diff_data.get('taux_changement')}%

PARAGRAPHES AJOUTÉS :
{ajoutes}

PARAGRAPHES SUPPRIMÉS :
{supprimes}

PARAGRAPHES MODIFIÉS :
{modifies}

Réponds UNIQUEMENT avec un JSON valide (aucun texte avant ou après), strictement dans ce format :
{{
  "resume_executif": "résumé des changements en 3 phrases",
  "changements_importants": [
    "changement 1",
    "changement 2",
    "changement 3",
    "changement 4",
    "changement 5"
  ],
  "interpretation": "ce que ces changements signifient pour le business",
  "alerte": "oui" ou "non",
  "score_importance": <entier de 1 à 10>
}}"""


def analyze(doc1_nom, doc2_nom, diff_data):
    """
    Analyse les différences entre deux rapports Word via Claude.

    Retourne un dict avec resume_executif, changements_importants,
    interpretation, alerte et score_importance.
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY manquant dans .env")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    prompt = _build_prompt(doc1_nom, doc2_nom, diff_data)

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )

        raw = message.content[0].text.strip()

        # Extrait le JSON si Claude a ajouté du texte autour
        start = raw.find("{")
        end = raw.rfind("}") + 1
        if start == -1 or end == 0:
            raise ValueError(f"Aucun JSON trouvé dans la réponse : {raw}")

        return json.loads(raw[start:end])

    except json.JSONDecodeError as e:
        print(f"Erreur parsing JSON Claude : {e}")
        raise
    except anthropic.APIError as e:
        print(f"Erreur API Anthropic : {e}")
        raise
    except Exception as e:
        print(f"Erreur inattendue dans analyze() : {e}")
        raise
