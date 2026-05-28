"""
claude_analyzer.py — Analyse de diff entre deux rapports Word via LLM multi-fournisseurs
"""

import json
import os
from datetime import datetime

from dotenv import load_dotenv
from llm_client import call_json, active_providers

load_dotenv()

_MAX_DOC_CHARS  = 50_000
_MAX_DIFF_CHARS = 20_000
_TOKEN_LIMIT    = 180_000


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

    diff_section = (
        f"PARAGRAPHES MODIFIÉS :\n{modifies}"
        if diff_only else
        f"PARAGRAPHES AJOUTÉS :\n{ajoutes}\n\n"
        f"PARAGRAPHES SUPPRIMÉS :\n{supprimes}\n\n"
        f"PARAGRAPHES MODIFIÉS :\n{modifies}"
    )

    return (
        "Tu es un analyste expert en rapports d'activité boursière BRVM. "
        "Analyse les différences entre deux rapports Word.\n\n"
        f"DOCUMENT 1 : {doc1_nom}\n"
        f"  - Paragraphes : {doc1_stats.get('nb_paragraphes')} | "
        f"Mots : {doc1_stats.get('nb_mots')} | "
        f"Caractères : {doc1_stats.get('nb_caracteres')}\n\n"
        f"DOCUMENT 2 : {doc2_nom}\n"
        f"  - Paragraphes : {doc2_stats.get('nb_paragraphes')} | "
        f"Mots : {doc2_stats.get('nb_mots')} | "
        f"Caractères : {doc2_stats.get('nb_caracteres')}\n\n"
        f"RÉSUMÉ : {diff_data.get('resume_changements')}\n"
        f"TAUX DE CHANGEMENT : {diff_data.get('taux_changement')}%\n\n"
        f"{diff_section}\n\n"
        "Réponds UNIQUEMENT avec un JSON valide (aucun texte avant ou après) :\n"
        '{"resume_executif": "résumé en 3 phrases", '
        '"changements_importants": ["ch1", "ch2", "ch3", "ch4", "ch5"], '
        '"interpretation": "signification business", '
        '"alerte": "oui" ou "non", '
        '"score_importance": <1-10>}'
    )


def analyze(doc1_nom, doc2_nom, diff_data):
    """
    Analyse les différences entre deux rapports Word via LLM (multi-fournisseurs).
    Retourne un dict avec : resume_executif, changements_importants,
    interpretation, alerte, score_importance.
    """
    providers = active_providers()
    if not providers:
        raise ValueError(
            "Aucune clé API LLM configurée. "
            "Définissez au moins ANTHROPIC_API_KEY, GEMINI_API_KEY, "
            "DEEPSEEK_API_KEY ou MISTRAL_API_KEY dans .env"
        )

    prompt = _build_prompt(doc1_nom, doc2_nom, diff_data)
    estimated_tokens = len(prompt) // 4
    print(f"[{_now()}] Fournisseurs disponibles : {providers}")
    print(f"[{_now()}] Prompt estimé à ~{estimated_tokens} tokens, envoi au LLM...")

    if estimated_tokens > _TOKEN_LIMIT:
        print(f"[{_now()}] Prompt trop long, réduction automatique...")
        prompt = _build_prompt(doc1_nom, doc2_nom, diff_data,
                               max_doc=_MAX_DOC_CHARS // 4,
                               max_diff=_MAX_DIFF_CHARS // 4)
        estimated_tokens = len(prompt) // 4
        print(f"[{_now()}] Après réduction : ~{estimated_tokens} tokens")

    try:
        result = call_json(prompt, max_tokens=1024)
        print(f"[{_now()}] Score d'importance : {result.get('score_importance')}/10")
        return result
    except Exception as e:
        print(f"[{_now()}] Erreur première tentative ({e}), retry diff uniquement...")
        fallback_prompt = _build_prompt(doc1_nom, doc2_nom, diff_data, diff_only=True)
        return call_json(fallback_prompt, max_tokens=1024)
