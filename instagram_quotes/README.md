# Instagram Quote Automation — The Shared Human Experience

Automatically generate and post minimalist inspirational quote images to Instagram daily.  
Quotes focus on universal human themes: impermanence, connection, parenting moments, seasonal nostalgia, and the quiet poetry of ordinary life.

---

## Features

| Feature | Details |
|---|---|
| Quote database | 55+ starter quotes; self-resets when exhausted |
| AI curation | Claude reads film knowledge and surfaces sentimental quotes |
| Image generation | PIL/Pillow — 8 colour palettes × 4 layout variants |
| AI captions | Claude writes engaging, human-feeling captions |
| Scheduling | Posts at configurable times daily |
| Session persistence | Avoids Instagram re-login on every run |
| Analytics | Saves media ID, like/comment counts per post |
| Archive | Moves posted images out of the working directory |
| A/B testing | Generate all layout variants for visual comparison |

---

## Quick Start

### 1. Clone / download the project

```bash
git clone <repo-url>
cd instagram_quotes
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

> Python 3.11+ recommended.

### 3. Configure credentials

```bash
cp .env.example .env
# Edit .env with your Instagram username/password and (optionally) Anthropic API key
```

> **Security note:** `.env` is in `.gitignore`. Never commit it.

### 4. (Optional) Add custom fonts

Drop `.ttf` or `.otf` font files into the `fonts/` directory.  
The system auto-detects them. Without custom fonts it falls back to system fonts (DejaVu, Liberation) or PIL's built-in font.

### 5. Preview without posting

```bash
python main.py --preview
```

This generates an image in `generated_images/` and prints the caption — no Instagram login needed.

### 6. Post once

```bash
python main.py --once
```

### 7. Run the daily scheduler

```bash
python main.py
```

The scheduler blocks and posts at the times configured in `config.py` → `POSTING_TIMES` (default: `09:00` and `18:00`).

---

## Command Reference

```
python main.py                          # Start scheduler (runs forever)
python main.py --once                   # Post once and exit
python main.py --preview                # Local preview — no Instagram
python main.py --curate                 # AI-curate quotes from a film
python main.py --curate --movie "Her" --theme "loneliness/connection"
python main.py --ab-test                # Generate all 4 layout variants
python main.py --once --layout split    # Force a specific layout
```

---

## Configuration

All settings live in `config.py`. Key options:

```python
POSTING_TIMES = ["09:00", "18:00"]     # 24-h format, local time
IMAGE_SIZE    = (1080, 1080)           # or (1080, 1350) for portrait
COLOR_PALETTES = [...]                  # 8 muted palettes — add your own
MAX_HASHTAGS  = 20
```

---

## Quote Database

`quotes_database.json` ships with 55 starter quotes.  
Each quote has:

```json
{
  "id": "q001",
  "text": "...",
  "source": "Good Will Hunting",
  "type": "movie",
  "themes": ["impermanence", "connection"],
  "used": false,
  "used_date": null
}
```

**Types:** `movie`, `tv`, `attributed`, `original`  
**Themes:** `impermanence`, `connection`, `loneliness`, `bittersweet parenting`, `seasonal nostalgia`, `quiet after chaos`, `small rituals`, `universal tiny moments`, `time passing`, `purpose`

### Add quotes manually

Open `quotes_database.json` and append entries following the schema above.

### AI curation

```bash
python main.py --curate --movie "Manchester by the Sea" --theme "grief/fatherhood"
```

Requires `ANTHROPIC_API_KEY` in `.env`. New quotes are appended to the database automatically.

---

## Deploying 24/7

### Option A — systemd (Linux VPS)

```ini
# /etc/systemd/system/instagram-quotes.service
[Unit]
Description=Instagram Quote Automation
After=network.target

[Service]
User=youruser
WorkingDirectory=/path/to/instagram_quotes
ExecStart=/usr/bin/python3 main.py
Restart=always
RestartSec=60
EnvironmentFile=/path/to/instagram_quotes/.env

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now instagram-quotes
sudo journalctl -u instagram-quotes -f
```

### Option B — screen / tmux

```bash
screen -S instagram
python main.py
# Ctrl+A D to detach
```

### Option C — cron (post once at a fixed time)

```cron
0 9  * * * cd /path/to/instagram_quotes && python main.py --once >> logs/cron.log 2>&1
0 18 * * * cd /path/to/instagram_quotes && python main.py --once >> logs/cron.log 2>&1
```

### Option D — Cloud (Railway / Render / Fly.io)

1. Push code to a private GitHub repo (ensure `.env` is in `.gitignore`).
2. Add environment variables via the platform's dashboard.
3. Set the start command to `python main.py`.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `ChallengeRequired` on login | Log in manually via the Instagram app, then retry |
| `BadPassword` | Double-check `.env` credentials |
| Blurry / ugly text | Add a TTF font to `fonts/` |
| No AI captions | Set `ANTHROPIC_API_KEY` in `.env` |
| Rate limit errors | Increase `POST_DELAY_SECONDS` and `RETRY_BACKOFF_SECONDS` in `config.py` |
| Quote pool exhausted | The system auto-resets; or run `--curate` to add more |

---

## Project Structure

```
instagram_quotes/
├── main.py                 # Orchestrator & CLI
├── quote_generator.py      # Quote selection, AI curation, caption generation
├── image_creator.py        # PIL image generation (4 layouts, 8 palettes)
├── instagram_poster.py     # instagrapi login, post, analytics, archive
├── config.py               # All settings (edit here)
├── requirements.txt
├── quotes_database.json    # 55+ starter quotes
├── .env.example            # Credential template
├── fonts/                  # Drop .ttf files here
├── generated_images/       # Working directory for new images
├── archive/                # Posted images moved here
├── logs/                   # automation.log
└── analytics.json          # Auto-created — post history & stats
```

---

## License

MIT — use freely, credit appreciated.
