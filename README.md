# Media Corpus Analysis

This repository contains the code and aggregate results for a media-focused social-media corpus collected across January–March 2025 and January–March 2026. The project examines media-related keywords, token usage, slang, hashtags, and engagement differences between the two periods.

## Key results

The included aggregate summary reports 26,395 tweets, 23,375 unique users, and 433,395 analysed tokens. The corpus contains 13,242 tweets from Q1 2025 and 13,153 from Q1 2026. The statistical tests use likes as an engagement proxy because manual sentiment scores were not completed.

See [`results/summary.csv`](results/summary.csv) and [`results/statistical_tests.csv`](results/statistical_tests.csv) for the reported values.

## Project status

This repository contains code from a completed course research project.
The scripts are preserved primarily for documentation and portfolio purposes,
not as a production-ready package.

## Reproducing the analysis

The raw workbook is not committed because it contains tweet text, author metadata, identifiers, and engagement data. Place the private workbook at `data/raw/media_corpus.xlsx`, then run:

```bash
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python src/analyze_media_corpus.py \
  --input data/raw/media_corpus.xlsx \
  --output-dir results
```

The analysis script also creates `results/token_level_analysis.csv`. That file is ignored by Git because it contains tweet IDs and can be large. Review it locally rather than publishing it.

## Collecting new data

The collection script uses the `twitterapi.io` Advanced Search endpoint. Supply the API key through the command line when running locally; never place the key in a source file, notebook, README, or commit:

```bash
python src/fetch_media_corpus.py \
  --api-key YOUR_API_KEY \
  --output data/raw/media_corpus.xlsx
```

The script writes a checkpoint CSV alongside the workbook. Both raw files are ignored by Git.

## Privacy and data availability

Only aggregate results are included in this public repository. Raw posts and account-level fields remain private. The raw data can be regenerated with the collection script, subject to the source platform's terms, the API provider's terms, and applicable research requirements.

## License

No open-source license is included. Without a license, the code may be viewed but is not automatically granted reuse rights.
