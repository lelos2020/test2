"""
Instagram posting logic using instagrapi.
Handles login, session persistence, posting, and analytics.
"""

import json
import logging
import os
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from instagrapi import Client
from instagrapi.exceptions import (
    BadPassword,
    ChallengeRequired,
    LoginRequired,
    TwoFactorRequired,
)

import config

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

def _build_client() -> Client:
    cl = Client()
    cl.delay_range = [2, 5]     # random delay between actions to mimic human behaviour
    return cl


def load_session(cl: Client) -> bool:
    """Load a saved session; return True if successful."""
    if not os.path.exists(config.SESSION_FILE):
        return False
    try:
        cl.load_settings(config.SESSION_FILE)
        cl.login(config.INSTAGRAM_USERNAME, config.INSTAGRAM_PASSWORD)
        logger.info("Session loaded from %s", config.SESSION_FILE)
        return True
    except Exception as exc:
        logger.warning("Session load failed (%s) — will do fresh login.", exc)
        return False


def save_session(cl: Client) -> None:
    cl.dump_settings(config.SESSION_FILE)
    logger.info("Session saved to %s", config.SESSION_FILE)


def get_authenticated_client() -> Client:
    """Return an authenticated instagrapi Client, reusing a saved session if possible."""
    if not config.INSTAGRAM_USERNAME or not config.INSTAGRAM_PASSWORD:
        raise ValueError("INSTAGRAM_USERNAME and INSTAGRAM_PASSWORD must be set in .env")

    cl = _build_client()

    if load_session(cl):
        return cl

    # Fresh login
    try:
        cl.login(config.INSTAGRAM_USERNAME, config.INSTAGRAM_PASSWORD)
        save_session(cl)
        return cl
    except BadPassword:
        raise RuntimeError("Instagram login failed: bad password.")
    except TwoFactorRequired:
        code = input("Enter your Instagram 2FA code: ").strip()
        cl.login(config.INSTAGRAM_USERNAME, config.INSTAGRAM_PASSWORD, verification_code=code)
        save_session(cl)
        return cl
    except ChallengeRequired:
        logger.error(
            "Instagram challenge required. Log in manually once via the app, then retry."
        )
        raise


# ---------------------------------------------------------------------------
# Posting
# ---------------------------------------------------------------------------

def post_photo(
    cl: Client,
    image_path: str,
    caption: str,
    retries: int = config.MAX_RETRIES,
) -> Optional[str]:
    """
    Upload a photo to Instagram.
    Returns the media ID on success, None on failure after all retries.
    """
    for attempt in range(1, retries + 1):
        try:
            logger.info("Posting image (attempt %d/%d): %s", attempt, retries, image_path)
            time.sleep(config.POST_DELAY_SECONDS)
            media = cl.photo_upload(image_path, caption)
            media_id = str(media.id)
            logger.info("Posted successfully — media ID: %s", media_id)
            return media_id
        except LoginRequired:
            logger.warning("Session expired — re-authenticating.")
            cl.login(config.INSTAGRAM_USERNAME, config.INSTAGRAM_PASSWORD)
            save_session(cl)
        except Exception as exc:
            wait = config.RETRY_BACKOFF_SECONDS * attempt
            logger.error("Post attempt %d failed: %s. Retrying in %ds.", attempt, exc, wait)
            time.sleep(wait)

    logger.error("All %d post attempts failed for %s.", retries, image_path)
    return None


# ---------------------------------------------------------------------------
# Analytics
# ---------------------------------------------------------------------------

def fetch_media_stats(cl: Client, media_id: str) -> dict:
    """Fetch like/comment/save counts for a posted media."""
    try:
        info = cl.media_info(media_id)
        return {
            "like_count": info.like_count,
            "comment_count": info.comment_count,
        }
    except Exception as exc:
        logger.warning("Could not fetch stats for %s: %s", media_id, exc)
        return {}


def record_analytics(media_id: str, quote: dict, image_path: str, stats: Optional[dict] = None) -> None:
    """Append a posting record to the analytics JSON file."""
    if not config.TRACK_ANALYTICS:
        return

    record = {
        "timestamp": datetime.now().isoformat(),
        "media_id": media_id,
        "quote_id": quote.get("id"),
        "quote_source": quote.get("source"),
        "image_path": image_path,
        "stats": stats or {},
    }

    data = []
    if os.path.exists(config.ANALYTICS_FILE):
        with open(config.ANALYTICS_FILE, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                data = []

    data.append(record)
    with open(config.ANALYTICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    logger.info("Analytics recorded for media %s.", media_id)


# ---------------------------------------------------------------------------
# Archive
# ---------------------------------------------------------------------------

def archive_image(image_path: str) -> str:
    """Move a posted image to the archive directory."""
    os.makedirs(config.ARCHIVE_DIR, exist_ok=True)
    dest = os.path.join(config.ARCHIVE_DIR, Path(image_path).name)
    shutil.move(image_path, dest)
    logger.info("Archived image to %s", dest)
    return dest
