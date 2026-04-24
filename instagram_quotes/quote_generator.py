"""
Quote selection, AI-assisted curation, and caption generation.
"""

import json
import logging
import os
import random
from datetime import date
from typing import Optional

import anthropic

import config

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def load_database() -> dict:
    if not os.path.exists(config.QUOTES_DB_PATH):
        raise FileNotFoundError(f"Quote database not found: {config.QUOTES_DB_PATH}")
    with open(config.QUOTES_DB_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_database(db: dict) -> None:
    with open(config.QUOTES_DB_PATH, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Quote selection
# ---------------------------------------------------------------------------

def pick_quote(theme: Optional[str] = None, reset_if_exhausted: bool = True) -> dict:
    """
    Return an unused quote from the database.
    Optionally filter by theme.  Resets all quotes if the pool is exhausted.
    """
    db = load_database()
    quotes = db["quotes"]

    unused = [q for q in quotes if not q.get("used", False)]
    if theme:
        themed = [q for q in unused if theme.lower() in [t.lower() for t in q.get("themes", [])]]
        pool = themed if themed else unused
    else:
        pool = unused

    if not pool:
        if reset_if_exhausted:
            logger.info("Quote pool exhausted — resetting all used flags.")
            for q in quotes:
                q["used"] = False
                q["used_date"] = None
            save_database(db)
            pool = quotes
        else:
            raise RuntimeError("No unused quotes available.")

    chosen = random.choice(pool)

    # Mark as used
    for q in quotes:
        if q["id"] == chosen["id"]:
            q["used"] = True
            q["used_date"] = date.today().isoformat()
    save_database(db)

    logger.info("Selected quote %s from '%s'", chosen["id"], chosen.get("source", "Original"))
    return chosen


# ---------------------------------------------------------------------------
# AI-assisted curation  (Claude via Anthropic SDK)
# ---------------------------------------------------------------------------

def ai_curate_from_script(movie_title: str, theme: str = "fatherhood/nostalgia", count: int = 5) -> list[dict]:
    """
    Ask Claude to surface sentimental quotes from a movie script/description.
    Returns a list of new quote dicts ready to append to the database.
    """
    if not config.ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY not set — skipping AI curation.")
        return []

    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)

    prompt = (
        f"You are a film scholar and emotional-intelligence curator. "
        f"Read your knowledge of the film '{movie_title}' and give me {count} sentimental quotes "
        f"that speak to the theme of '{theme}'. "
        f"For each quote return a JSON object with keys: text, source, themes (list of strings). "
        f"Themes should be chosen from: impermanence, connection, loneliness, bittersweet parenting, "
        f"seasonal nostalgia, quiet after chaos, small rituals, universal tiny moments, time passing, purpose. "
        f"Return ONLY a JSON array, no commentary."
    )

    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = message.content[0].text.strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    try:
        curated = json.loads(raw)
    except json.JSONDecodeError as exc:
        logger.error("AI curation JSON parse error: %s", exc)
        return []

    db = load_database()
    existing_texts = {q["text"].strip().lower() for q in db["quotes"]}
    new_quotes = []
    for item in curated:
        text = item.get("text", "").strip()
        if not text or text.lower() in existing_texts:
            continue
        new_id = f"ai{len(db['quotes']) + len(new_quotes) + 1:04d}"
        quote = {
            "id": new_id,
            "text": text,
            "source": item.get("source", movie_title),
            "type": "movie",
            "themes": item.get("themes", []),
            "used": False,
            "used_date": None,
        }
        new_quotes.append(quote)

    if new_quotes:
        db["quotes"].extend(new_quotes)
        db["metadata"]["total_quotes"] = len(db["quotes"])
        db["metadata"]["last_updated"] = date.today().isoformat()
        save_database(db)
        logger.info("AI curation added %d new quotes from '%s'.", len(new_quotes), movie_title)

    return new_quotes


# ---------------------------------------------------------------------------
# Caption generation
# ---------------------------------------------------------------------------

def generate_caption(quote: dict, use_ai: bool = True) -> str:
    """
    Build the Instagram caption: an engaging expansion of the quote + hashtags.
    Falls back to a template if AI is unavailable.
    """
    if use_ai and config.ANTHROPIC_API_KEY:
        caption = _ai_caption(quote)
    else:
        caption = _template_caption(quote)

    hashtags = _build_hashtags(quote)
    return f"{caption}\n\n{hashtags}"


def _ai_caption(quote: dict) -> str:
    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    prompt = (
        f"Write a short, emotionally resonant Instagram caption (2-3 sentences) that expands on this quote. "
        f"The caption should feel warm, relatable, and human — like a friend sharing something true. "
        f"Do not restate the quote verbatim. End with an open question that invites reflection.\n\n"
        f"Quote: \"{quote['text']}\"\n"
        f"Source: {quote.get('source', 'Original')}\n"
        f"Themes: {', '.join(quote.get('themes', []))}"
    )
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=256,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text.strip()


def _template_caption(quote: dict) -> str:
    themes = quote.get("themes", ["universal tiny moments"])
    theme = themes[0].replace("_", " ") if themes else "the small moments"
    source = quote.get("source", "Original")

    templates = [
        f"There's something about {theme} that makes us realize we've never been alone in this. "
        f"{'— ' + source if source != 'Original' else ''}\n\nWhat moment this week made you feel most human?",

        f"Some truths land quietly. This one stopped me in my tracks. "
        f"{'(from ' + source + ')' if source != 'Original' else ''}\n\nWhat does this stir up in you?",

        f"If you've ever felt this, you already know. And if you haven't — keep living. "
        f"{'— ' + source if source != 'Original' else ''}\n\nWhat's a moment you'll never forget?",
    ]
    return random.choice(templates)


def _build_hashtags(quote: dict) -> str:
    theme_tag_map = {
        "impermanence": ["#Impermanence", "#NothingLasts", "#TransientMoments"],
        "connection": ["#HumanConnection", "#WeAreNotAlone", "#TrueConnection"],
        "loneliness": ["#BeautifullyAlone", "#InnerWorld", "#SolitudeIsGold"],
        "bittersweet parenting": ["#ParentingMoments", "#KidsGrowUp", "#BitersweetLove"],
        "seasonal nostalgia": ["#SeasonalNostalgia", "#AutumnFeels", "#SummerMemories"],
        "quiet after chaos": ["#AfterTheParty", "#PeacefulMoments", "#StillnessIsBeautiful"],
        "small rituals": ["#SmallRituals", "#MorningRoutine", "#EverydayMagic"],
        "universal tiny moments": ["#TinyMoments", "#EverydayPoetry", "#SliceOfLife"],
        "time passing": ["#TimeFlies", "#GrowingOlder", "#LifeIsShort"],
        "purpose": ["#FindYourPurpose", "#LivingIntentionally", "#MeaningfulLife"],
    }

    tags = list(config.BASE_HASHTAGS)
    for theme in quote.get("themes", []):
        extra = theme_tag_map.get(theme.lower(), [])
        tags.extend(extra)

    tags = list(dict.fromkeys(tags))  # deduplicate, preserve order
    return " ".join(tags[: config.MAX_HASHTAGS])
