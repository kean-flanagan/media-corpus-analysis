# Data dictionary and privacy notes

## Raw corpus fields

The private workbook contains post-level fields including:

- `Tweet ID`, date, period, keyword, language, and source
- raw and cleaned text
- author ID, author name, and username
- verification and bot-candidate indicators
- follower counts, post counts, likes, reposts, replies, quotes, views, and bookmarks
- hashtags and optional manual annotation fields

The checkpoint CSV contains the same information in a machine-oriented naming convention.

## Public result files

The `results/` directory contains aggregate outputs only:

- `summary.csv`: corpus totals and high-level rates
- `keyword_frequency.csv`: keyword counts by period/year
- `hashtag_frequency.csv`: aggregate hashtag frequencies
- `top_tokens_by_period.csv`: top content tokens by period
- `statistical_tests.csv`: period comparisons by keyword
- `within_year_comparison.csv`: within-year comparisons

The generated `token_level_analysis.csv` is not public. Although it contains token-level rather than full-post text, it retains tweet IDs and is therefore ignored by Git.

## Interpretation note

The current statistical tests use `Likes` as an engagement proxy. The `Sentiment` annotation field is present in the workbook but was not completed in the supplied analysis output.
