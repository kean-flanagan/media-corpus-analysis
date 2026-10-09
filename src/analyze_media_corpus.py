#!/usr/bin/env python3
"""
analyze_media_corpus.py
───────────────────────
Runs NLP and statistical analysis on the media_corpus.xlsx workbook
produced by fetch_media_corpus.py.

Outputs (written to ./results/ by default):
    token_level_analysis.csv         — every token with POS, lang, and frequency
    keyword_frequency.csv            — keyword counts per period and year
    top_tokens_by_period.csv         — top-50 content tokens per period
    statistical_tests.csv            — t-test, Cohen's d, Mann-Whitney U per keyword
    within_year_comparison.csv       — month-over-month within Q1-2025 and Q1-2026
    hashtag_frequency.csv            — top hashtags by period
    summary.csv                      — high-level counts and rates

Usage:
    python3 analyze_media_corpus.py

    # Or point at a different workbook:
    python3 analyze_media_corpus.py --input my_corpus.xlsx

Requirements:
    pip install pandas openpyxl spacy scipy
    python -m spacy download en_core_web_sm

Notes:
    - Analysis runs on Text (clean) — the emoji-stripped column.
    - Tokens that are stopwords, URLs, @-mentions, or punctuation are dropped
      before frequency analysis, matching the proposal's cleaning goals.
    - Slang detection uses a small starter list; expand SLANG_TOKENS for
      better coverage (the proposal mentions filtering slang for cleaner data).
    - Statistical tests operate on *like-count* as a numeric engagement proxy
      when manual sentiment annotations are absent. Once you fill in the
      Sentiment column with numeric scores, swap the test_col variable to
      'sentiment_score'.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
import spacy
from scipy import stats

# ── Config ────────────────────────────────────────────────────────────────────

INPUT_XLSX   = Path("media_corpus.xlsx")
OUTPUT_DIR   = Path("results")
DEDUP_BY_ID  = True
TOP_N_TOKENS = 50

# Media keywords from the proposal (used as a second filter pass)
MEDIA_KEYWORDS = [
    "movie", "film", "cinema",
    "album", "song", "artist",
    "meme", "viral", "trend",
]

# excludes viral (as part of statistical observation)
MEDIA_KEYWORDS = [
    "movie", "film", "cinema",
    "album", "song", "artist",
    "meme", "trend",
]

# Starter slang list — expand as needed
SLANG_TOKENS = {
    "lol", "lmao", "omg", "smh", "tbh", "imo", "fr", "ngl", "idk",
    "bruh", "bro", "sis", "fam", "lowkey", "highkey", "slay", "lit",
    "bussin", "no cap", "based", "mid", "goated", "w", "l", "rn",
}

URL_RE   = re.compile(r"https?://\S+")
MENTION_RE = re.compile(r"@\w+")


# ── Helpers ───────────────────────────────────────────────────────────────────

def cohen_d(a, b) -> float:
    """Pooled Cohen's d for two independent samples."""
    n1, n2 = len(a), len(b)
    if n1 < 2 or n2 < 2:
        return float("nan")
    s = (((n1 - 1) * a.std(ddof=1) ** 2 + (n2 - 1) * b.std(ddof=1) ** 2) / (n1 + n2 - 2)) ** 0.5
    return float((a.mean() - b.mean()) / s) if s else float("nan")


def clean_for_nlp(text: str) -> str:
    text = URL_RE.sub("", text)
    text = MENTION_RE.sub("", text)
    return text.strip()


def is_slang(token_text: str) -> bool:
    return token_text.lower() in SLANG_TOKENS


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=str(INPUT_XLSX))
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR),
                        help="Directory for generated CSV files (default: results)")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    if not input_path.exists():
        sys.exit(f"[ERROR] File not found: {input_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Load ──────────────────────────────────────────────────────────────────
    df = pd.read_excel(input_path, dtype={"Tweet ID": str})
    df = df[df["Text (clean)"].notna()].copy()

    if DEDUP_BY_ID and "Tweet ID" in df.columns:
        df = df.drop_duplicates(subset=["Tweet ID"]).copy()

    # Normalise column names defensively
    df.columns = [c.strip() for c in df.columns]

    required = ["Tweet ID", "Period", "Year Period", "Keyword", "Text (clean)", "Likes"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        sys.exit(f"[ERROR] Missing expected columns: {missing}")

    df["Likes"]    = pd.to_numeric(df["Likes"], errors="coerce").fillna(0)
    df["Retweets"] = pd.to_numeric(df.get("Retweets", 0), errors="coerce").fillna(0)
    df["Views"]    = pd.to_numeric(df.get("Views", 0), errors="coerce").fillna(0)

    # Use 'Sentiment' column if manually filled with numeric scores
    if "Sentiment" in df.columns:
        df["sentiment_score"] = pd.to_numeric(df["Sentiment"], errors="coerce")
        test_col = "sentiment_score"
    else:
        test_col = "Likes"

    periods = sorted(df["Period"].dropna().unique().tolist())
    print(f"Loaded {len(df):,} tweets across periods: {periods}")

    # ── NLP: tokenise with spaCy ──────────────────────────────────────────────
    print("Loading spaCy model …")
    try:
        nlp = spacy.load("en_core_web_sm", disable=["ner"])
    except OSError:
        sys.exit(
            "[ERROR] spaCy model not found.\n"
            "Run:  python -m spacy download en_core_web_sm"
        )

    print("Tokenising tweets …")
    token_records = []
    for _, row in df.iterrows():
        text = clean_for_nlp(str(row["Text (clean)"]))
        doc  = nlp(text)
        for tok in doc:
            if tok.is_space or tok.is_punct or tok.like_url or tok.like_email:
                continue
            if tok.is_stop:
                continue
            token_records.append({
                "Tweet ID":    row["Tweet ID"],
                "Period":      row["Period"],
                "Year Period": row["Year Period"],
                "Keyword":     row["Keyword"],
                "token":       tok.text,
                "lemma":       tok.lemma_.lower(),
                "pos":         tok.pos_,
                "tag":         tok.tag_,
                "is_slang":    is_slang(tok.text),
                "is_media_kw": tok.lemma_.lower() in MEDIA_KEYWORDS,
            })

    token_df = pd.DataFrame(token_records)
    token_df.to_csv(output_dir / "token_level_analysis.csv", index=False)
    print(f"  → token_level_analysis.csv  ({len(token_df):,} tokens)")

    # ── Keyword frequency by period ───────────────────────────────────────────
    kw_freq = (
        df.groupby(["Period", "Year Period", "Keyword"])
        .size()
        .reset_index(name="tweet_count")
        .sort_values(["Period", "Keyword"])
    )
    kw_freq.to_csv(output_dir / "keyword_frequency.csv", index=False)
    print(f"  → keyword_frequency.csv")

    # ── Top content tokens per period ─────────────────────────────────────────
    content_pos = {"NOUN", "VERB", "ADJ", "ADV", "PROPN"}
    content_tokens = token_df[token_df["pos"].isin(content_pos) & ~token_df["is_slang"]]
    top_tokens = (
        content_tokens.groupby(["Period", "lemma", "pos"])
        .size()
        .reset_index(name="count")
        .sort_values(["Period", "count"], ascending=[True, False])
        .groupby("Period")
        .head(TOP_N_TOKENS)
        .reset_index(drop=True)
    )
    top_tokens.to_csv(output_dir / "top_tokens_by_period.csv", index=False)
    print(f"  → top_tokens_by_period.csv  (top {TOP_N_TOKENS} per period)")

    # ── Hashtag frequency ─────────────────────────────────────────────────────
    if "Hashtags" in df.columns:
        ht_rows = []
        for _, row in df.iterrows():
            raw = str(row.get("Hashtags", ""))
            for tag in [t.strip().lower() for t in raw.split(",") if t.strip()]:
                ht_rows.append({"Period": row["Period"], "Year Period": row["Year Period"], "hashtag": tag})
        ht_df = pd.DataFrame(ht_rows)
        if not ht_df.empty:
            ht_freq = (
                ht_df.groupby(["Period", "hashtag"])
                .size()
                .reset_index(name="count")
                .sort_values(["Period", "count"], ascending=[True, False])
                .groupby("Period")
                .head(20)
                .reset_index(drop=True)
            )
            ht_freq.to_csv(output_dir / "hashtag_frequency.csv", index=False)
            print(f"  → hashtag_frequency.csv")

    # ── Statistical tests: Q1-2025 vs Q1-2026 per keyword ────────────────────
    print("\nRunning statistical tests …")
    test_records = []
    for keyword in MEDIA_KEYWORDS:
        sub = df[df["Keyword"] == keyword]
        g25 = sub[sub["Year Period"] == "Q1-2025"][test_col].dropna()
        g26 = sub[sub["Year Period"] == "Q1-2026"][test_col].dropna()

        if len(g25) < 5 or len(g26) < 5:
            test_records.append({
                "keyword": keyword, "n_2025": len(g25), "n_2026": len(g26),
                "mean_2025": round(g25.mean(), 2) if len(g25) else None,
                "mean_2026": round(g26.mean(), 2) if len(g26) else None,
                "ttest_t": None, "ttest_p": None,
                "cohens_d": None,
                "mannwhitney_u": None, "mannwhitney_p": None,
                "test_column": test_col, "note": "insufficient data (n<5)",
            })
            continue

        t_stat, t_p  = stats.ttest_ind(g25, g26, equal_var=False)
        u_stat, u_p  = stats.mannwhitneyu(g25, g26, alternative="two-sided")
        d            = cohen_d(g25, g26)

        test_records.append({
            "keyword":       keyword,
            "n_2025":        len(g25),
            "n_2026":        len(g26),
            "mean_2025":     round(g25.mean(), 2),
            "mean_2026":     round(g26.mean(), 2),
            "ttest_t":       round(t_stat, 4),
            "ttest_p":       round(t_p, 4),
            "cohens_d":      round(d, 4),
            "mannwhitney_u": round(u_stat, 2),
            "mannwhitney_p": round(u_p, 4),
            "test_column":   test_col,
            "note":          "",
        })

    test_df = pd.DataFrame(test_records)
    test_df.to_csv(output_dir / "statistical_tests.csv", index=False)
    print(f"  → statistical_tests.csv")

    # ── Within-year month comparison (recency bias check) ────────────────────
    month_order = {
        "2025-Jan": 1, "2025-Feb": 2, "2025-Mar": 3,
        "2026-Jan": 4, "2026-Feb": 5, "2026-Mar": 6,
    }
    within_records = []
    for year_label in ("Q1-2025", "Q1-2026"):
        sub = df[df["Year Period"] == year_label].copy()
        sub["month_n"] = sub["Period"].map(month_order)
        for keyword in MEDIA_KEYWORDS:
            ksub = sub[sub["Keyword"] == keyword]
            months = sorted(ksub["Period"].unique(), key=lambda p: month_order.get(p, 99))
            for i in range(len(months) - 1):
                m1, m2 = months[i], months[i + 1]
                g1 = ksub[ksub["Period"] == m1][test_col].dropna()
                g2 = ksub[ksub["Period"] == m2][test_col].dropna()
                if len(g1) < 5 or len(g2) < 5:
                    continue
                t_stat, t_p = stats.ttest_ind(g1, g2, equal_var=False)
                within_records.append({
                    "year_period":  year_label,
                    "keyword":      keyword,
                    "month_a":      m1,
                    "month_b":      m2,
                    "n_a":          len(g1),
                    "n_b":          len(g2),
                    "mean_a":       round(g1.mean(), 2),
                    "mean_b":       round(g2.mean(), 2),
                    "ttest_t":      round(t_stat, 4),
                    "ttest_p":      round(t_p, 4),
                    "test_column":  test_col,
                })

    within_df = pd.DataFrame(within_records)
    within_df.to_csv(output_dir / "within_year_comparison.csv", index=False)
    print(f"  → within_year_comparison.csv")

    # ── High-level summary ────────────────────────────────────────────────────
    bot_col = "Bot Candidate" if "Bot Candidate" in df.columns else None
    ver_col = "Blue Verified"  if "Blue Verified"  in df.columns else None

    summary_data = {
        "metric": [
            "total_tweets",
            "tweets_Q1_2025",
            "tweets_Q1_2026",
            "unique_users",
            "bot_candidates",
            "blue_verified",
            "avg_likes_2025",
            "avg_likes_2026",
            "slang_tokens",
            "total_tokens_analysed",
            "test_column_used",
        ],
        "value": [
            len(df),
            int((df["Year Period"] == "Q1-2025").sum()),
            int((df["Year Period"] == "Q1-2026").sum()),
            df["Username"].nunique() if "Username" in df.columns else "n/a",
            int(df[bot_col].sum()) if bot_col else "n/a",
            int(df[ver_col].sum()) if ver_col else "n/a",
            round(df[df["Year Period"] == "Q1-2025"]["Likes"].mean(), 2),
            round(df[df["Year Period"] == "Q1-2026"]["Likes"].mean(), 2),
            int(token_df["is_slang"].sum()),
            len(token_df),
            test_col,
        ],
    }
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_csv(output_dir / "summary.csv", index=False)
    print(f"  → summary.csv")

    # ── Console summary ───────────────────────────────────────────────────────
    print("\n=== Corpus Summary ===")
    print(summary_df.to_string(index=False))
    print("\n=== Statistical Tests (cross-year) ===")
    print(test_df[["keyword", "n_2025", "n_2026", "mean_2025", "mean_2026",
                    "ttest_p", "cohens_d", "mannwhitney_p"]].to_string(index=False))
    print(f"\nAll outputs written to: {output_dir.resolve()}")
    if test_col == "Likes":
        print(
            "\nNote: statistical tests ran on 'Likes' (engagement proxy) "
            "because the Sentiment column is not yet filled in. "
            "Fill it with numeric scores and re-run to test sentiment directly."
        )


if __name__ == "__main__":
    main()
