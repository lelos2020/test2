"""
Central configuration for the Instagram Quote Automation system.
All user-tunable settings live here. Credentials come from .env via python-dotenv.
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Instagram credentials  (set in .env, never hard-code)
# ---------------------------------------------------------------------------
INSTAGRAM_USERNAME = os.getenv("INSTAGRAM_USERNAME", "")
INSTAGRAM_PASSWORD = os.getenv("INSTAGRAM_PASSWORD", "")

# ---------------------------------------------------------------------------
# Anthropic / Claude API  (for AI-assisted quote curation)
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUOTES_DB_PATH = os.path.join(BASE_DIR, "quotes_database.json")
GENERATED_IMAGES_DIR = os.path.join(BASE_DIR, "generated_images")
ARCHIVE_DIR = os.path.join(BASE_DIR, "archive")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
SESSION_FILE = os.path.join(BASE_DIR, "instagram_session.json")

# ---------------------------------------------------------------------------
# Scheduling
# ---------------------------------------------------------------------------
# 24-hour format strings; multiple times = multiple daily posts
POSTING_TIMES = ["09:00", "18:00"]

# ---------------------------------------------------------------------------
# Image settings
# ---------------------------------------------------------------------------
IMAGE_SIZE = (1080, 1080)          # square; use (1080, 1350) for portrait
FONT_DIR = os.path.join(BASE_DIR, "fonts")

# Muted, warm palette cycles per post
COLOR_PALETTES = [
    {"bg": "#F5F0E8", "text": "#3D2B1F", "accent": "#C4956A"},   # warm parchment
    {"bg": "#E8EDE8", "text": "#2C3E2D", "accent": "#7A9E7E"},   # sage green
    {"bg": "#EDE8E3", "text": "#4A3728", "accent": "#C17F5A"},   # terracotta
    {"bg": "#E3E8ED", "text": "#1E2D3D", "accent": "#5A7C9E"},   # dusty blue
    {"bg": "#EDE3E8", "text": "#3D1E2D", "accent": "#9E5A7C"},   # dusty mauve
    {"bg": "#FAFAF7", "text": "#2D2D2D", "accent": "#8B8B6B"},   # clean cream
    {"bg": "#F0E8D8", "text": "#3C2415", "accent": "#B8860B"},   # golden wheat
    {"bg": "#D8E8E0", "text": "#1A3028", "accent": "#4E8B6F"},   # deep sage
]

# ---------------------------------------------------------------------------
# Hashtags
# ---------------------------------------------------------------------------
BASE_HASHTAGS = [
    "#SharedHumanExperience",
    "#RelatableQuotes",
    "#TheseSmallMoments",
    "#QuietObservations",
    "#SliceOfLife",
    "#HumanConnection",
    "#Impermanence",
    "#MindfulLiving",
    "#EveryDayMoments",
    "#Nostalgia",
    "#TimeFlies",
    "#ParentingMoments",
    "#FeelingsAreUniversal",
    "#InspirationalQuotes",
    "#WordsThatHit",
    "#FilmQuotes",
    "#CinematicWisdom",
    "#PoetryOfLife",
    "#EmotionalIntelligence",
    "#BeautifullyHuman",
]

# Max hashtags per post (Instagram limit is 30)
MAX_HASHTAGS = 20

# ---------------------------------------------------------------------------
# Retry / rate-limit settings
# ---------------------------------------------------------------------------
POST_DELAY_SECONDS = 5          # delay between API calls during posting
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 30      # base wait between retries

# ---------------------------------------------------------------------------
# A/B testing
# ---------------------------------------------------------------------------
# Layouts: "centered", "top_heavy", "bottom_heavy", "split"
LAYOUT_OPTIONS = ["centered", "top_heavy", "bottom_heavy", "split"]

# ---------------------------------------------------------------------------
# Analytics
# ---------------------------------------------------------------------------
ANALYTICS_FILE = os.path.join(BASE_DIR, "analytics.json")
TRACK_ANALYTICS = True
