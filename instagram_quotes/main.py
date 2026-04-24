"""
Orchestrator for the Instagram Quote Automation system.

Usage:
    python main.py                  # run the scheduler (blocks forever)
    python main.py --once           # post once immediately and exit
    python main.py --curate         # AI-curate quotes from a movie and exit
    python main.py --preview        # generate an image locally without posting
    python main.py --ab-test        # generate all layout variants for today's quote
"""

import argparse
import logging
import os
import sys
import time

import schedule

import config
from image_creator import create_image
from instagram_poster import (
    archive_image,
    fetch_media_stats,
    get_authenticated_client,
    post_photo,
    record_analytics,
)
from quote_generator import ai_curate_from_script, generate_caption, pick_quote

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

os.makedirs(config.LOGS_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(config.LOGS_DIR, "automation.log"), encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Core posting job
# ---------------------------------------------------------------------------

def run_posting_job(preview_only: bool = False, layout: str = None) -> bool:
    """
    Full pipeline: pick quote → generate image → post to Instagram.
    If preview_only=True, skips the Instagram step.
    Returns True on success.
    """
    logger.info("=== Starting posting job (preview=%s) ===", preview_only)

    # 1. Pick a quote
    try:
        quote = pick_quote()
    except Exception as exc:
        logger.error("Quote selection failed: %s", exc)
        return False

    logger.info("Quote: %s", quote["text"][:80])

    # 2. Generate image
    try:
        image_path = create_image(quote, layout=layout)
    except Exception as exc:
        logger.error("Image creation failed: %s", exc)
        return False

    if preview_only:
        logger.info("Preview mode — image saved at: %s", image_path)
        caption = generate_caption(quote, use_ai=bool(config.ANTHROPIC_API_KEY))
        print("\n--- CAPTION PREVIEW ---")
        print(caption)
        print("-----------------------\n")
        return True

    # 3. Build caption
    try:
        caption = generate_caption(quote, use_ai=bool(config.ANTHROPIC_API_KEY))
    except Exception as exc:
        logger.error("Caption generation failed: %s", exc)
        return False

    # 4. Post to Instagram
    try:
        cl = get_authenticated_client()
    except Exception as exc:
        logger.error("Instagram authentication failed: %s", exc)
        return False

    media_id = post_photo(cl, image_path, caption)

    if not media_id:
        logger.error("Posting failed — image kept at %s for manual retry.", image_path)
        return False

    # 5. Analytics
    stats = fetch_media_stats(cl, media_id)
    record_analytics(media_id, quote, image_path, stats)

    # 6. Archive
    archive_image(image_path)

    logger.info("=== Job complete. Media ID: %s ===", media_id)
    return True


# ---------------------------------------------------------------------------
# A/B test helper
# ---------------------------------------------------------------------------

def run_ab_test() -> None:
    """Generate one image per layout variant so you can compare them."""
    quote = pick_quote()
    for layout in config.LAYOUT_OPTIONS:
        path = create_image(quote, layout=layout)
        logger.info("A/B variant '%s' → %s", layout, path)
    print("All layout variants generated in:", config.GENERATED_IMAGES_DIR)


# ---------------------------------------------------------------------------
# Scheduler
# ---------------------------------------------------------------------------

def setup_schedule() -> None:
    for posting_time in config.POSTING_TIMES:
        schedule.every().day.at(posting_time).do(run_posting_job)
        logger.info("Scheduled daily post at %s", posting_time)


def run_scheduler() -> None:
    setup_schedule()
    logger.info("Scheduler running. Press Ctrl+C to stop.")
    while True:
        schedule.run_pending()
        time.sleep(30)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Instagram Quote Automation")
    parser.add_argument("--once", action="store_true", help="Post once immediately and exit")
    parser.add_argument("--preview", action="store_true", help="Generate image locally without posting")
    parser.add_argument("--curate", action="store_true", help="AI-curate quotes from a film")
    parser.add_argument("--ab-test", action="store_true", help="Generate all layout variants")
    parser.add_argument("--movie", default="Good Will Hunting", help="Movie title for --curate")
    parser.add_argument("--theme", default="fatherhood/nostalgia", help="Theme for --curate")
    parser.add_argument("--layout", default=None, help="Force a specific layout: centered|top_heavy|bottom_heavy|split")
    args = parser.parse_args()

    if args.curate:
        new_quotes = ai_curate_from_script(args.movie, args.theme)
        print(f"Added {len(new_quotes)} new quotes from '{args.movie}'.")
        return

    if args.ab_test:
        run_ab_test()
        return

    if args.once or args.preview:
        success = run_posting_job(preview_only=args.preview, layout=args.layout)
        sys.exit(0 if success else 1)

    # Default: run the scheduler indefinitely
    run_scheduler()


if __name__ == "__main__":
    main()
