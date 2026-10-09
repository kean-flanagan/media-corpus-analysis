#!/usr/bin/env python3
"""
fetch_media_corpus.py
─────────────────────
Builds a media-sentiment corpus from X (Twitter) using twitterapi.io's
Advanced Search endpoint. Collects tweets matching media keywords across
six monthly windows (Jan–Mar 2025 and Jan–Mar 2026) for cross-year comparison.

Each month is collected as a separate batch and tagged with its period label
so you can compare within-year and across-year in the output workbook.

Usage:
    python3 fetch_media_corpus.py --api-key <YOUR_TWITTERAPI_IO_KEY>

Optional flags:
    --tweets-per-month   Target tweet count per month-keyword pair (default: 200)
    --output             Output filename            (default: media_corpus.xlsx)
    --lang               Twitter lang filter         (default: en)
    --exclude-verified   Flag verified/blue accounts in output (default: True)

Requirements:
    pip install requests pandas openpyxl

Notes:
    - Advanced Search returns ≤20 tweets per page; the script paginates
      using cursor until the target is hit or pages run out.
    - twitterapi.io explicitly warns against cursor-based pagination across
      large result sets; time-slicing (since_time/until_time) is used instead.
    - Retweets are excluded by default (-filter:retweets) to reduce noise.
    - Emojis are stripped in a post-processing step before export.
    - A is_bot_candidate flag is set when an account has >50 tweets/day
      on average since creation — a rough but useful heuristic.
"""

import argparse
import re
import sys
import time
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────

TWEETS_PER_MONTH_DEFAULT = 500
RANDOM_SEED              = 42
API_URL                  = "https://api.twitterapi.io/twitter/tweet/advanced_search"
PAGE_DELAY_S             = 0.3   # polite delay between paginated calls
BOT_TWEETS_PER_DAY       = 50    # heuristic threshold for bot flagging

# Media keywords from the proposal
MEDIA_KEYWORDS = [
    "movie", "film", "cinema",
    "album", "song", "artist",
    "meme", "viral", "trend",
]

# Monthly windows: (label, since_YYYY-MM-DD, until_YYYY-MM-DD)
MONTHLY_WINDOWS = [
    ("2025-Jan", "2025-01-01", "2025-01-31"),
    ("2025-Feb", "2025-02-01", "2025-02-28"),
    ("2025-Mar", "2025-03-01", "2025-03-31"),
    ("2026-Jan", "2026-01-01", "2026-01-31"),
    ("2026-Feb", "2026-02-01", "2026-02-28"),
    ("2026-Mar", "2026-03-01", "2026-03-31"),
]

# Aggregate year-period labels for cross-year comparison column
YEAR_PERIOD = {
    "2025-Jan": "Q1-2025", "2025-Feb": "Q1-2025", "2025-Mar": "Q1-2025",
    "2026-Jan": "Q1-2026", "2026-Feb": "Q1-2026", "2026-Mar": "Q1-2026",
}

# ── Colours ───────────────────────────────────────────────────────────────────

CLR_HEADER_BG  = "1DA1F2"
CLR_HEADER_FG  = "FFFFFF"
CLR_ROW_A      = "E3F2FD"
CLR_ROW_B      = "F5F5F5"
CLR_2025_BG    = "E8F5E9"
CLR_2026_BG    = "FFF9C4"
CLR_BOT_BG     = "FFCDD2"
CLR_ANNOT_BG   = "FFFDE7"

# ── Helpers ───────────────────────────────────────────────────────────────────

EMOJI_RE = re.compile(
    "["
    "\U0001F600-\U0001F64F"
    "\U0001F300-\U0001F5FF"
    "\U0001F680-\U0001F6FF"
    "\U0001F1E0-\U0001F1FF"
    "\U00002700-\U000027BF"
    "\U0000FE00-\U0000FE0F"
    "\U00002600-\U000026FF"
    "\U0001F900-\U0001F9FF"
    "\U00003030"
    "]+",
    flags=re.UNICODE,
)

def strip_emojis(text: str) -> str:
    text = EMOJI_RE.sub("", text)
    # Strip control characters and anything else openpyxl rejects
    text = re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', text)
    text = ''.join(c for c in text if c.isprintable() or c in ('\n', '\t'))
    return text.strip()


def sanitize_cell(val):
    """Sanitize any value before writing to an Excel cell."""
    if isinstance(val, str):
        return strip_emojis(val)
    return val


def is_bot_candidate(author: dict) -> bool:
    """Rough heuristic: avg posts/day since account creation."""
    created_raw = author.get("createdAt", "")
    statuses    = author.get("statusesCount", 0)
    if not created_raw or not statuses:
        return False
    try:
        created_dt = datetime.strptime(created_raw, "%a %b %d %H:%M:%S +0000 %Y").replace(
            tzinfo=timezone.utc
        )
        days = max((datetime.now(timezone.utc) - created_dt).days, 1)
        return (statuses / days) > BOT_TWEETS_PER_DAY
    except ValueError:
        return False


def build_query(keyword: str, since: str, until: str, lang: str) -> str:
    """
    Build an Advanced Search query string.
    since/until format expected by twitterapi.io: YYYY-MM-DD_HH:MM:SS_UTC
    """
    since_ts = f"{since}_00:00:00_UTC"
    until_ts  = f"{until}_23:59:59_UTC"
    # Exclude retweets, filter to language
    return (
        f'"{keyword}" -filter:retweets lang:{lang} '
        f"since:{since_ts} until:{until_ts}"
    )


def fetch_page(query: str, cursor: str, api_key: str) -> dict:
    headers = {"X-API-Key": api_key}
    params  = {"query": query, "queryType": "Latest", "cursor": cursor}
    try:
        resp = requests.get(API_URL, headers=headers, params=params, timeout=30)
    except requests.RequestException as e:
        print(f"  [WARN] Request error: {e}", flush=True)
        return {}

    if resp.status_code == 401:
        sys.exit("[ERROR] Invalid API key.")

    if resp.status_code == 429:
        print("  [Rate limited] Waiting 10 s …", flush=True)
        time.sleep(10)
        return fetch_page(query, cursor, api_key)

    if not resp.ok:
        print(f"  [WARN] HTTP {resp.status_code}: {resp.text[:200]}", flush=True)
        return {}

    return resp.json()


def collect_tweets(
    keyword: str,
    label: str,
    since: str,
    until: str,
    target: int,
    lang: str,
    api_key: str,
) -> list[dict]:
    query   = build_query(keyword, since, until, lang)
    cursor  = ""
    rows    = []
    seen    = set()

    while len(rows) < target:
        data = fetch_page(query, cursor, api_key)
        if not data:
            break

        for tweet in data.get("tweets", []):
            tid = tweet.get("id")
            if not tid or tid in seen:
                continue
            seen.add(tid)

            text   = tweet.get("text", "")
            author = tweet.get("author", {})

            rows.append({
                "tweet_id":           tid,
                "period":             label,
                "year_period":        YEAR_PERIOD.get(label, ""),
                "keyword":            keyword,
                "text_raw":           text,
                "text_clean":         strip_emojis(text),
                "created_at":         str(tweet.get("createdAt", ""))[:10],
                "lang":               tweet.get("lang", ""),
                "source":             tweet.get("source", ""),
                "is_reply":           tweet.get("isReply", False),
                "author_id":          author.get("id", ""),
                "author_name":        author.get("name", ""),
                "username":           author.get("userName", ""),
                "is_blue_verified":   author.get("isBlueVerified", False),
                "is_bot_candidate":   is_bot_candidate(author),
                "author_followers":   author.get("followers", ""),
                "author_statuses":    author.get("statusesCount", ""),
                "likes":              tweet.get("likeCount", ""),
                "retweets":           tweet.get("retweetCount", ""),
                "replies":            tweet.get("replyCount", ""),
                "quotes":             tweet.get("quoteCount", ""),
                "views":              tweet.get("viewCount", ""),
                "bookmarks":          tweet.get("bookmarkCount", ""),
                "hashtags":           ", ".join(
                    h.get("text", "") for h in
                    tweet.get("entities", {}).get("hashtags", [])
                ),
                "sentiment":          "",  # annotation column
                "topic_medium":       "",  # annotation: movie / music / meme / other
                "notes":              "",  # annotation
            })

            if len(rows) >= target:
                break

        has_next = data.get("has_next_page", False)
        cursor   = data.get("next_cursor", "")
        if not has_next or not cursor:
            break

        time.sleep(PAGE_DELAY_S)

    return rows


# ── Excel helpers ─────────────────────────────────────────────────────────────

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def header_font():
    return Font(bold=True, color=CLR_HEADER_FG, name="Arial", size=10)


def apply_header(ws, headers: list[str]):
    fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    for col, label in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.font      = header_font()
        cell.fill      = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin_border()
    ws.row_dimensions[1].height = 32


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Fetch media-sentiment tweets via twitterapi.io Advanced Search and export to Excel."
    )
    parser.add_argument("--api-key",           required=True)
    parser.add_argument("--tweets-per-month",  type=int, default=TWEETS_PER_MONTH_DEFAULT)
    parser.add_argument("--output",            default="media_corpus.xlsx")
    parser.add_argument("--lang",              default="en")
    parser.add_argument("--exclude-verified",  action="store_true", default=False,
                        help="Skip blue-verified accounts entirely")
    args = parser.parse_args()

    all_rows: list[dict] = []
    total_combos = len(MONTHLY_WINDOWS) * len(MEDIA_KEYWORDS)
    combo_n      = 0

    for label, since, until in MONTHLY_WINDOWS:
        for keyword in MEDIA_KEYWORDS:
            combo_n += 1
            print(
                f"[{combo_n}/{total_combos}] {label}  keyword='{keyword}'  "
                f"target={args.tweets_per_month} …",
                flush=True,
            )
            rows = collect_tweets(
                keyword, label, since, until,
                args.tweets_per_month, args.lang, args.api_key,
            )
            if args.exclude_verified:
                rows = [r for r in rows if not r["is_blue_verified"]]
            print(f"    → collected {len(rows)} tweets", flush=True)
            all_rows.extend(rows)

    print(f"\nTotal tweets collected: {len(all_rows)}")

    # ── Deduplicate across keyword queries ────────────────────────────────────
    seen_ids: set[str] = set()
    deduped: list[dict] = []
    for r in all_rows:
        if r["tweet_id"] not in seen_ids:
            seen_ids.add(r["tweet_id"])
            deduped.append(r)
    print(f"After deduplication: {len(deduped)} tweets")

    # ── Checkpoint: save raw data before Excel write ──────────────────────────
    checkpoint_path = args.output.replace(".xlsx", "_checkpoint.csv")
    pd.DataFrame(deduped).to_csv(checkpoint_path, index=False, encoding="utf-8-sig")
    print(f"Checkpoint saved → {checkpoint_path}")

    # ── Write Excel ───────────────────────────────────────────────────────────
    print(f"\nWriting {args.output} …")
    wb = Workbook()

    # ── Sheet 1: Corpus ───────────────────────────────────────────────────────
    ws1            = wb.active
    ws1.title      = "Corpus"
    ws1.freeze_panes = "A2"

    headers = [
        "Tweet ID", "Period", "Year Period", "Keyword",
        "Text (raw)", "Text (clean)", "Date", "Lang", "Source",
        "Is Reply",
        "Author ID", "Author Name", "Username",
        "Blue Verified", "Bot Candidate",
        "Followers", "Statuses Count",
        "Likes", "Retweets", "Replies", "Quotes", "Views", "Bookmarks",
        "Hashtags",
        "Sentiment ✏️", "Topic Medium ✏️", "Notes ✏️",
    ]
    apply_header(ws1, headers)

    col_widths = [
        20, 12, 12, 12,
        80, 80, 12, 8, 22,
        10,
        20, 22, 18,
        14, 14,
        12, 14,
        8, 10, 8, 8, 10, 10,
        28,
        18, 20, 30,
    ]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    fill_2025 = PatternFill("solid", fgColor=CLR_2025_BG)
    fill_2026 = PatternFill("solid", fgColor=CLR_2026_BG)
    fill_bot  = PatternFill("solid", fgColor=CLR_BOT_BG)
    fill_ann  = PatternFill("solid", fgColor=CLR_ANNOT_BG)

    for r_idx, row in enumerate(deduped, 2):
        values = [
            row["tweet_id"], row["period"], row["year_period"], row["keyword"],
            row["text_raw"], row["text_clean"], row["created_at"], row["lang"], row["source"],
            row["is_reply"],
            row["author_id"], row["author_name"], row["username"],
            row["is_blue_verified"], row["is_bot_candidate"],
            row["author_followers"], row["author_statuses"],
            row["likes"], row["retweets"], row["replies"], row["quotes"], row["views"], row["bookmarks"],
            row["hashtags"],
            row["sentiment"], row["topic_medium"], row["notes"],
        ]

        is_bot  = row["is_bot_candidate"]
        is_2025 = row["period"].startswith("2025")
        base_fill = fill_bot if is_bot else (fill_2025 if is_2025 else fill_2026)

        for c_idx, val in enumerate(values, 1):
            cell           = ws1.cell(row=r_idx, column=c_idx, value=sanitize_cell(val))
            cell.font      = Font(name="Arial", size=9)
            cell.border    = thin_border()
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=(c_idx in (5, 6)),
            )
            cell.fill = fill_ann if c_idx >= 25 else base_fill

        ws1.row_dimensions[r_idx].height = 55

    ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # ── Sheet 2: Monthly Summary ──────────────────────────────────────────────
    ws2       = wb.create_sheet("Monthly Summary")
    ws2.freeze_panes = "A2"

    df = pd.DataFrame(deduped)

    sum_headers = ["Period", "Year Period", "Keyword", "Tweet Count",
                   "Bot Candidates", "Blue Verified", "Avg Likes", "Avg Retweets", "Avg Views"]
    apply_header(ws2, sum_headers)
    for i, w in enumerate([14, 12, 14, 14, 14, 14, 12, 14, 12], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    fill_a = PatternFill("solid", fgColor=CLR_ROW_A)
    fill_b = PatternFill("solid", fgColor=CLR_ROW_B)

    r_idx = 2
    for label, since, until in MONTHLY_WINDOWS:
        yp = YEAR_PERIOD.get(label, "")
        for keyword in MEDIA_KEYWORDS:
            subset = df[(df["period"] == label) & (df["keyword"] == keyword)]
            n = len(subset)
            if n == 0:
                continue
            avg = lambda col: round(pd.to_numeric(subset[col], errors="coerce").mean(), 1)
            row_vals = [
                label, yp, keyword, n,
                int(subset["is_bot_candidate"].sum()),
                int(subset["is_blue_verified"].sum()),
                avg("likes"), avg("retweets"), avg("views"),
            ]
            row_fill = fill_a if r_idx % 2 == 0 else fill_b
            for c_idx, val in enumerate(row_vals, 1):
                cell        = ws2.cell(row=r_idx, column=c_idx, value=val)
                cell.font   = Font(name="Arial", size=9)
                cell.border = thin_border()
                cell.fill   = row_fill
                cell.alignment = Alignment(vertical="center")
            ws2.row_dimensions[r_idx].height = 18
            r_idx += 1

    ws2.auto_filter.ref = f"A1:{get_column_letter(len(sum_headers))}1"

    # ── Sheet 3: Annotation Guide ─────────────────────────────────────────────
    ws3 = wb.create_sheet("Annotation Guide")
    ws3.column_dimensions["A"].width = 26
    ws3.column_dimensions["B"].width = 60

    guide = [
        ("Field",             "Description"),
        ("Text (raw)",        "Original tweet text including emojis"),
        ("Text (clean)",      "Emoji-stripped version — use for NLP"),
        ("Period",            "Month label: 2025-Jan … 2026-Mar"),
        ("Year Period",       "Aggregate label: Q1-2025 or Q1-2026"),
        ("Keyword",           "The media keyword that surfaced this tweet"),
        ("Is Reply",          "True = the tweet is a reply to another tweet"),
        ("Bot Candidate",     f"True = account posts >{BOT_TWEETS_PER_DAY} tweets/day on avg (heuristic — verify manually)"),
        ("Blue Verified",     "True = account has Twitter Blue / verified badge"),
        ("",                  ""),
        ("── Annotation Columns ──", ""),
        ("Sentiment ✏️",      "positive / negative / neutral / mixed"),
        ("Topic Medium ✏️",   "movie  /  music  /  meme  /  other"),
        ("Notes ✏️",          "Free-form observations (slang, sarcasm, context, etc.)"),
        ("",                  ""),
        ("── Statistical Tests ──", ""),
        ("t-test",            "Compare mean sentiment scores between Q1-2025 and Q1-2026"),
        ("Cohen's d",         "Effect size for the t-test; supplements the p-value"),
        ("Mann-Whitney U",    "Non-parametric alternative if sentiment distributions are skewed"),
        ("Within-year",       "Compare months inside the same year to check recency bias"),
    ]

    hdr_fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    for i, (label, desc) in enumerate(guide, 1):
        c1 = ws3.cell(row=i, column=1, value=label)
        c2 = ws3.cell(row=i, column=2, value=desc)
        for c in (c1, c2):
            c.font      = Font(name="Arial", size=10, bold=(i == 1))
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border    = thin_border()
            if i == 1:
                c.fill = hdr_fill
                c.font = header_font()
        ws3.row_dimensions[i].height = 20

    wb.save(args.output)
    print(f"\nSaved → {args.output}")
    print(
        f"\nCorpus ready. Colour key:\n"
        f"  Green rows  = 2025 tweets\n"
        f"  Yellow rows = 2026 tweets\n"
        f"  Red rows    = bot candidates (review before including in analysis)\n"
        f"  Yellow cells (columns 25-27) = annotation columns\n"
    )


if __name__ == "__main__":
    main()
