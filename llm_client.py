"""
llm_client.py — Client LLM unifié multi-fournisseurs
=====================================================
Fournisseurs : Anthropic (Claude) · Google (Gemini) · DeepSeek · Mistral
Stratégie    : cascade automatique — essaie dans l'ordre, fallback sur erreur ou clé absente

Variables d'environnement requises (au moins une) :
  ANTHROPIC_API_KEY   — Claude Sonnet (primaire)
  GEMINI_API_KEY      — Gemini 1.5 Pro (fallback 1)
  DEEPSEEK_API_KEY    — DeepSeek Chat (fallback 2, idéal JSON structuré)
  MISTRAL_API_KEY     — Mistral Large (fallback 3, bon support français)
"""

import json
import logging
import os
import re
from typing import Optional

logger = logging.getLogger(__name__)

# ── Modèles par fournisseur ───────────────────────────────────────────────────

MODELS = {
    "anthropic": "claude-sonnet-4-20250514",
    "gemini":    "gemini-1.5-pro",
    "deepseek":  "deepseek-chat",
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
    import google.generativeai as genai
    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
    model = genai.GenerativeModel(
        MODELS["gemini"],
        system_instruction=system or "Tu es un expert financier BRVM spécialisé en analyse boursière UEMOA.",
    )
    response = model.generate_content(
        prompt,
        generation_config=genai.GenerationConfig(max_output_tokens=max_tokens),
    )
    return response.text.strip()


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
