"""
llm_client.py — Client LLM unifié multi-fournisseurs
=====================================================
Fournisseurs : Anthropic (Claude) · Google (Gemini) · DeepSeek · Mistral
Stratégie    : cascade automatique — essaie dans l'ordre, fallback sur erreur ou clé absente

Variables d'environnement requises (au moins une) :
  ANTHROPIC_API_KEY   — Claude Sonnet (primaire)
  GEMINI_API_KEY      — Gemini 2.5 Flash (fallback 1)
  DEEPSEEK_API_KEY    — DeepSeek V4 Flash (fallback 2, idéal JSON structuré)
  MISTRAL_API_KEY     — Mistral Large (fallback 3, bon support français)
"""

import json
import logging
import os
import re
from typing import Optional

logger = logging.getLogger(__name__)

# ── Modèles par fournisseur ───────────────────────────────────────────────────
# MàJ 2026-07 :
#   - gemini-1.5-pro  → arrêté (404).
#   - gemini-2.5-flash → « no longer available to new users » (404).
#     Remplacé par gemini-3.5-flash-lite (GA, bon marché, idéal extraction JSON).
#   - deepseek-chat   → alias retiré le 24/07/2026. Remplacé par deepseek-v4-flash.

MODELS = {
    "anthropic": "claude-sonnet-4-20250514",
    "gemini":    "gemini-3.5-flash-lite",
    "deepseek":  "deepseek-v4-flash",
    "mistral":   "mistral-large-latest",
}

# Ordre de priorité (cascade) : Gemini → DeepSeek → Mistral → Claude (dernier recours)
PROVIDER_ORDER = ["gemini", "deepseek", "mistral", "anthropic"]

_ENV_KEYS = {
    "anthropic": "ANTHROPIC_API_KEY",
    "gemini":    "GEMINI_API_KEY",
    "deepseek":  "DEEPSEEK_API_KEY",
    "mistral":   "MISTRAL_API_KEY",
}


# ── Détection des clés disponibles ───────────────────────────────────────────

def has_key(provider: str) -> bool:
    val = os.getenv(_ENV_KEYS.get(provider, ""), "")
    return bool(val and val.strip())


def active_providers() -> list:
    """Retourne la liste ordonnée des fournisseurs ayant une clé API valide."""
    return [p for p in PROVIDER_ORDER if has_key(p)]


# ── Appelants par fournisseur ─────────────────────────────────────────────────

def _call_anthropic(prompt: str, max_tokens: int, system: Optional[str]) -> str:
    import anthropic
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    kwargs = {
        "model":     MODELS["anthropic"],
        "max_tokens": max_tokens,
        "messages":  [{"role": "user", "content": prompt}],
    }
    if system:
        kwargs["system"] = system
    msg = client.messages.create(**kwargs)
    return msg.content[0].text.strip()


def _call_gemini(prompt: str, max_tokens: int, system: Optional[str]) -> str:
    # Nouveau SDK : google-genai (l'ancien google.generativeai n'est plus maintenu).
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
    config = types.GenerateContentConfig(
        system_instruction=system
        or "Tu es un expert financier BRVM spécialisé en analyse boursière UEMOA.",
        max_output_tokens=max_tokens,
        # Les modèles Gemini 3.x activent le "thinking" par défaut, dont les tokens
        # sont décomptés de max_output_tokens → réponse vide si le budget est modeste.
        # Pour l'extraction/classification à haut volume, on force le niveau minimal.
        thinking_config=types.ThinkingConfig(thinking_level="minimal"),
    )
    response = client.models.generate_content(
        model=MODELS["gemini"],
        contents=prompt,
        config=config,
    )
    return (response.text or "").strip()


def _call_deepseek(prompt: str, max_tokens: int, system: Optional[str]) -> str:
    from openai import OpenAI
    client = OpenAI(
        api_key=os.getenv("DEEPSEEK_API_KEY"),
        base_url="https://api.deepseek.com",
    )
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    resp = client.chat.completions.create(
        model=MODELS["deepseek"],
        max_tokens=max_tokens,
        messages=messages,
    )
    return resp.choices[0].message.content.strip()


def _call_mistral(prompt: str, max_tokens: int, system: Optional[str]) -> str:
    from mistralai import Mistral
    client = Mistral(api_key=os.getenv("MISTRAL_API_KEY"))
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    resp = client.chat.complete(
        model=MODELS["mistral"],
        max_tokens=max_tokens,
        messages=messages,
    )
    return resp.choices[0].message.content.strip()


_CALLERS = {
    "anthropic": _call_anthropic,
    "gemini":    _call_gemini,
    "deepseek":  _call_deepseek,
    "mistral":   _call_mistral,
}


# ── API publique ──────────────────────────────────────────────────────────────

def call(
    prompt: str,
    max_tokens: int = 1024,
    system: Optional[str] = None,
    provider: str = "auto",
) -> str:
    """
    Appelle un LLM et retourne la réponse texte brute.

    provider="auto" : essaie chaque fournisseur dans PROVIDER_ORDER
                      jusqu'au premier succès.
    provider="anthropic"|"gemini"|"deepseek"|"mistral" : force ce fournisseur.
    """
    order = PROVIDER_ORDER if provider == "auto" else [provider]
    errors = []

    for p in order:
        if not has_key(p):
            logger.debug("[LLM] %s : clé absente, skip", p)
            continue
        try:
            logger.info("[LLM] Appel %s (model=%s, max_tokens=%d)", p, MODELS[p], max_tokens)
            result = _CALLERS[p](prompt, max_tokens, system)
            # Une réponse vide (0 caractère utile) NE doit pas être considérée comme
            # un succès : certains modèles (raisonnement) renvoient un contenu vide
            # quand le budget de tokens est consommé par la réflexion. On force alors
            # le passage au fournisseur suivant au lieu de renvoyer "".
            if not result or not result.strip():
                raise RuntimeError("réponse vide (0 caractère)")
            logger.info("[LLM] Succès %s — %d chars", p, len(result))
            print(f"  [LLM] ✓ {p} ({MODELS[p]}) — {len(result)} chars")
            return result
        except Exception as exc:
            err_msg = f"{p}: {exc}"
            logger.warning("[LLM] Échec %s : %s", p, exc)
            print(f"  [LLM] ✗ {p} : {exc} → tentative suivante...")
            errors.append(err_msg)

    raise RuntimeError(
        f"Tous les fournisseurs LLM ont échoué ({', '.join(p for p in order if has_key(p))}). "
        f"Détails : {errors}"
    )


def call_json(
    prompt: str,
    max_tokens: int = 2048,
    system: Optional[str] = None,
    provider: str = "auto",
) -> "dict | list":
    """
    Comme call() mais parse et retourne le JSON extrait de la réponse.
    Extrait le premier tableau [] ou objet {} trouvé (ignore les blocs ```json```).
    Lève ValueError si aucun JSON valide n'est trouvé.
    """
    raw = call(prompt, max_tokens, system, provider)
    cleaned = re.sub(r"```(?:json)?\s*", "", raw).strip().rstrip("`").strip()

    # Tableau en priorité (cas le plus fréquent dans extractor)
    s_arr = cleaned.find("[")
    s_obj = cleaned.find("{")

    if s_arr != -1 and (s_obj == -1 or s_arr < s_obj):
        e = cleaned.rfind("]") + 1
        if e > 0:
            return json.loads(cleaned[s_arr:e])

    if s_obj != -1:
        e = cleaned.rfind("}") + 1
        if e > 0:
            return json.loads(cleaned[s_obj:e])

    raise ValueError(f"Aucun JSON valide dans la réponse LLM : {raw[:300]}")
