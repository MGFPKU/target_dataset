# 📄 Climate Target Candidate Extractor

A rule-based Python script to extract candidate sentences containing climate targets from Chinese and English policy documents (HTML and PDF).

The tool is designed for high-recall screening: it identifies sentences that are likely to contain quantitative or time-bound targets, which can then be reviewed manually or used for downstream processing.

## 🔍 Overview

This script processes collections of policy documents and:

- extracts raw text from HTML and PDF files
- optionally applies OCR for scanned PDFs
- splits text into sentences (language-specific rules)
- applies regex-based filters to identify candidate target sentences
- exports results to CSV or Excel

The output is a structured table of candidate sentences for further analysis.

## ⚠️ Important note

This is a heuristic, rule-based extraction tool, not a full target identification system.

- ✔️ Designed for broad coverage (recall)
- ❗ May include false positives
- ❗ May miss some valid targets

👉 Output should be reviewed manually or refined further.

## 📂 Supported inputs

**Languages:**

- Chinese (`--lang cn`)
- English (`--lang en`)

**Formats:**

- HTML (`.htm`, `.html`)
- PDF (`.pdf`)

## ⚙️ Installation

1. Clone the repository

   ```bash
   git clone https://github.com/MGFPKU/target_dataset.git
   cd target_dataset
