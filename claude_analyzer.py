import os
import json
import anthropic
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

_MAX_DOC_CHARS = 50_000   # paragraphes_ajoutes et paragraphes_supprimes
_MAX_DIFF_CHARS = 20_000  # paragraphes_modifies
_TOKEN_LIMIT = 180_000


def _now():
    return datetime.now().strftime("%Y-%m")


def _truncate(text, max_chars):
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n[... contenu tronqué ...]"


def _build_prompt(doc1_nom, doc2_nom, diff_data, diff_only=False,
                  max_doc=_MAX_DOC_CHARS, max_diff=_MAX_DIFF_CHARS):
    doc1_stats = diff_data.get("doc1_stats", {})
    doc2_stats = diff_data.get("doc2_stats", {})

    ajoutes = _truncate(
        "\n".join(f"  - {p}" for p in diff_data.get("paragraphes_ajoutes", [])) or "  Aucun",
        max_doc,
    )
    supprimes = _truncate(
        "\n".join(f"  - {p}" for p in diff_data.get("paragraphes_supprimes", [])) or "  Aucun",
        max_doc,
    )
    modifies = _truncate(
        "\n".join(
            f"  - AVANT: {m['avant']}\n    APRÈS: {m['apres']} (similarité: {m['similarite']})"
            for m in diff_data.get("paragraphes_modifies", [])
        ) or "  Aucun",
        max_diff,
    )

    if diff_only:
        diff_section = f"PARAGRAPHES MODIFIÉS :\n{modifies}"
    else:
        diff_section = (
            f"PARAGRAPHES AJOUTÉS :\n{ajoutes}\n\n"
            f"PARAGRAPHES SUPPRIMÉS :\n{supprimes}\n\n"
            f"PARAGRAPHES MODIFIÉS :\n{modifies}"
        )

    return (
        f"Tu es un analyste expert en rapports d'activité. Analyse les différences entre deux rapports Word.\n\n"
        f"DOCUMENT 1 : {doc1_nom}\n"
        f"  - Paragraphes : {doc1_stats.get('nb_paragraphes')} | Mots : {doc1_stats.get('nb_mots')} | Caractères : {doc1_stats.get('nb_caracteres')}\n\n"
        f"DOCUMENT 2 : {doc2_nom}\n"
        f"  - Paragraphes : {doc2_stats.get('nb_paragraphes')} | Mots : {doc2_stats.get('nb_mots')} | Caractères : {doc2_stats.get('nb_caracteres')}\n\n"
        f"RÉSUMÉ : {diff_data.get('resume_changements')}\n"
        f"TAUX DE CHANGEMENT : {diff_data.get('taux_changement')}%\n\n"
        f"{diff_section}\n\n"
        f"Réponds UNIQUEMENT avec un JSON valide (aucun texte avant ou après) :\n"
        f'{{"resume_executif": "résumé en 3 phrases", "changements_importants": ["ch1", "ch2", "ch3", "ch4", "ch5"], "interpretation": "signification business", "alerte": "oui" ou "non", "score_importance": <1-10>}}'
    )


def _parse_response(message):
    raw = message.content[0].text.strip()
    start = raw.find("{")
    end = raw.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError(f"Aucun JSON trouvé dans la réponse : {raw}")
    return json.loads(raw[start:end])


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

    # Vérification préventive de la taille
    estimated_tokens = len(prompt) // 4
    print(f"[{_now()}] Prompt estimé à ~{estimated_tokens} tokens, envoi à Claude...")

    if estimated_tokens > _TOKEN_LIMIT:
        print(f"[{_now()}] Prompt trop long, réduction automatique...")
        prompt = _build_prompt(doc1_nom, doc2_nom, diff_data,
                               max_doc=_MAX_DOC_CHARS // 4,
                               max_diff=_MAX_DIFF_CHARS // 4)
        estimated_tokens = len(prompt) // 4
        print(f"[{_now()}] Après réduction : ~{estimated_tokens} tokens, envoi à Claude...")

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        return _parse_response(message)

    except anthropic.BadRequestError as e:
        if getattr(e, "status_code", None) == 400 or "too long" in str(e).lower():
            print(f"[{_now()}] Retry avec diff uniquement...")
            fallback_prompt = _build_prompt(doc1_nom, doc2_nom, diff_data, diff_only=True)
            message = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1024,
                messages=[{"role": "user", "content": fallback_prompt}],
            )
            return _parse_response(message)
        raise
    except json.JSONDecodeError as e:
        print(f"Erreur parsing JSON Claude : {e}")
        raise
    except anthropic.APIError as e:
        print(f"Erreur API Anthropic : {e}")
        raise
    except Exception as e:
        print(f"Erreur inattendue dans analyze() : {e}")
        raise
