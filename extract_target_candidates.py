#!/usr/bin/env python3
"""
Rule-based candidate sentence extractor for climate targets in policy documents.

This script extracts candidate target-related sentences from Chinese and English
policy documents in HTML or PDF format. It uses language-specific regex patterns
to identify sentences that are likely to contain quantitative targets or target
formulations. For PDFs, an optional OCR fallback is available for scanned files.

The script is designed as a candidate sentence screening tool rather than a full
target extraction pipeline. Output should therefore be reviewed manually or used
as input for downstream processing.

Example usage:
    # Chinese HTML
    python extract_target_candidates.py \
        --lang cn --source html \
        --input ./data/html_cn \
        --output ./outputs/cn_html.xlsx

    # Chinese PDF with OCR fallback
    python extract_target_candidates.py \
        --lang cn --source pdf --ocr --tesseract-lang chi_sim \
        --input ./data/pdf_cn \
        --output ./outputs/cn_pdf.xlsx

    # English HTML
    python extract_target_candidates.py \
        --lang en --source html \
        --input ./data/html_en \
        --output ./outputs/en_html.xlsx

    # English PDF
    python extract_target_candidates.py \
        --lang en --source pdf \
        --input ./data/pdf_en \
        --output ./outputs/en_pdf.csv
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd
import pdfplumber
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------
# Regex patterns and language-specific rules
# ---------------------------------------------------------------------

CN_TARGET_PATTERN = re.compile(
    r"(到\d{4}年底|到\d{4}年|\d{4}[—\-]\d{4}年|与\d{4}年相比|比\d{4}年|"
    r"“十[一二三四五]五”期间|展望\d{4}年|\d{4}年前)"
)

EN_TARGET_PATTERN = re.compile(
    r"(by\s+\d{4}|target(s)?\s+to\s+reduce|aim(s)?\s+to\s+achieve|"
    r"compared\s+to\s+\d{4}|from\s+\d{4}|net\s+zero\s+by\s+\d{4}|"
    r"increase\s+by\s+\d{1,3}%|reduce\s+emissions\s+by\s+\d{1,3}%)",
    re.IGNORECASE,
)

DEFAULT_CN_SENTENCE_SPLIT = r"(?<=[。！？；])\s*"
DEFAULT_EN_SENTENCE_SPLIT = r"(?<=[.!?])\s+"

CN_EXCLUDES = {
    "exclude_start": re.compile(r"^(ariafocus|【|附件|专栏|表)"),
    "exclude_full_date": re.compile(r"\d{4}年\d{1,2}月\d{1,2}日"),
    "exclude_symbols": re.compile(r"[{\《表]"),
}

EN_EXCLUDES = {
    "exclude_start": re.compile(
        r"^(Appendix|Table|Figure|Reference|Note|\[|\*|\(|•)",
        re.IGNORECASE,
    ),
    "exclude_full_date": re.compile(
        r"(\d{1,2}(st|nd|rd|th)?\s+"
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}|"
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+"
        r"\d{1,2},\s+\d{4})",
        re.IGNORECASE,
    ),
}


# ---------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------

def iter_files(input_dir: Path, source: str, recursive: bool) -> Iterable[Path]:
    """
    Yield input files matching the requested source type.

    For HTML, both .htm and .html are supported.
    """
    if source == "pdf":
        patterns = ["**/*.pdf"] if recursive else ["*.pdf"]
    else:
        patterns = ["**/*.htm", "**/*.html"] if recursive else ["*.htm", "*.html"]

    seen = set()
    for pattern in patterns:
        for path in input_dir.glob(pattern):
            if path.is_file() and path not in seen:
                seen.add(path)
                yield path


def normalize_whitespace(text: str) -> str:
    """Collapse repeated whitespace and trim leading/trailing spaces."""
    return re.sub(r"\s+", " ", text).strip()


def remove_page_number_artifacts(text: str) -> str:
    """
    Remove common page number artifacts from extracted PDF text.

    Example:
        — 21 —
        - 21 -
    """
    return re.sub(r"[—\-]\s*\d+\s*[—\-]", " ", text)


def sentence_splitter(lang: str) -> re.Pattern:
    """Return the sentence-splitting regex for the requested language."""
    split_pattern = DEFAULT_CN_SENTENCE_SPLIT if lang == "cn" else DEFAULT_EN_SENTENCE_SPLIT
    return re.compile(split_pattern)


def get_rules(lang: str) -> Tuple[re.Pattern, Dict[str, re.Pattern], re.Pattern]:
    """Return target pattern, exclusion rules, and sentence splitter for the selected language."""
    if lang == "cn":
        return CN_TARGET_PATTERN, CN_EXCLUDES, sentence_splitter("cn")
    return EN_TARGET_PATTERN, EN_EXCLUDES, sentence_splitter("en")


def passes_filters(
    sentence: str,
    target_pattern: re.Pattern,
    excludes: Dict[str, re.Pattern],
    min_len: int,
) -> bool:
    """
    Check whether a sentence qualifies as a candidate target sentence.

    The filtering logic is intentionally simple and heuristic:
    - minimum length requirement
    - must contain at least one digit
    - must match a target-related regex pattern
    - must not match exclusion patterns
    """
    s = sentence.strip()

    if len(s) < min_len:
        return False

    # Require at least one digit to reduce non-target false positives.
    if not re.search(r"\d+", s):
        return False

    if not target_pattern.search(s):
        return False

    if "exclude_start" in excludes and excludes["exclude_start"].match(s):
        return False

    if "exclude_full_date" in excludes and excludes["exclude_full_date"].search(s):
        return False

    if "exclude_symbols" in excludes and excludes["exclude_symbols"].search(s):
        return False

    return True


def build_row(
    document_name: str,
    sentence: str,
    lang: str,
    source_type: str,
    extraction_method: str,
) -> Dict[str, str]:
    """Create one output row."""
    return {
        "Document": document_name,
        "Sentence": sentence.strip(),
        "Language": lang,
        "Source_Type": source_type,
        "Extraction_Method": extraction_method,
    }


# ---------------------------------------------------------------------
# HTML extraction
# ---------------------------------------------------------------------

def extract_from_html_file(
    html_path: Path,
    target_pattern: re.Pattern,
    excludes: Dict[str, re.Pattern],
    splitter: re.Pattern,
    min_len: int,
    lang: str,
) -> List[Dict[str, str]]:
    """
    Extract candidate target sentences from a single HTML file.
    """
    data = html_path.read_bytes()

    # Prefer lxml when available, but fall back to the built-in parser.
    try:
        soup = BeautifulSoup(data, "lxml")
    except Exception:
        soup = BeautifulSoup(data, "html.parser")

    text = soup.get_text(" ", strip=True)
    text = normalize_whitespace(text)
    sentences = splitter.split(text)

    rows = []
    for sentence in sentences:
        if passes_filters(sentence, target_pattern, excludes, min_len):
            rows.append(
                build_row(
                    document_name=html_path.name,
                    sentence=sentence,
                    lang=lang,
                    source_type="html",
                    extraction_method="html_text",
                )
            )

    return rows


# ---------------------------------------------------------------------
# PDF extraction (+ optional OCR fallback)
# ---------------------------------------------------------------------

def extract_text_pdfplumber(pdf_path: Path) -> str:
    """
    Extract text from a PDF using pdfplumber.
    """
    full_text = ""
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
    return full_text


def extract_text_ocr(pdf_path: Path, tesseract_lang: str) -> str:
    """
    Extract text from a scanned PDF using OCR.

    Note:
        This requires optional dependencies:
        - pdf2image
        - pytesseract

        In many environments, pdf2image also requires Poppler to be installed.
    """
    from pdf2image import convert_from_path
    import pytesseract

    full_text = ""
    images = convert_from_path(str(pdf_path))
    for image in images:
        full_text += pytesseract.image_to_string(image, lang=tesseract_lang) + "\n"
    return full_text


def extract_from_pdf_file(
    pdf_path: Path,
    target_pattern: re.Pattern,
    excludes: Dict[str, re.Pattern],
    splitter: re.Pattern,
    min_len: int,
    enable_ocr: bool,
    tesseract_lang: str,
    lang: str,
) -> List[Dict[str, str]]:
    """
    Extract candidate target sentences from a single PDF file.

    Workflow:
    1. Try normal text extraction with pdfplumber.
    2. If no text is found and OCR is enabled, try OCR.
    3. Clean and split text into sentences.
    4. Apply heuristic filters.
    """
    print(f"Processing PDF: {pdf_path.name}")

    extraction_method = "pdf_text"
    try:
        text = extract_text_pdfplumber(pdf_path)
    except Exception as exc:
        print(f"  pdfplumber failed for {pdf_path.name}: {exc}")
        text = ""

    if enable_ocr and not text.strip():
        print(f"  No extractable PDF text found. Trying OCR for {pdf_path.name}...")
        extraction_method = "ocr"
        try:
            text = extract_text_ocr(pdf_path, tesseract_lang=tesseract_lang)
        except Exception as exc:
            print(f"  OCR failed for {pdf_path.name}: {exc}")
            return []

    if not text.strip():
        print(f"  No usable text extracted from {pdf_path.name}.")
        return []

    text = remove_page_number_artifacts(text.replace("\n", " "))
    text = normalize_whitespace(text)
    sentences = splitter.split(text)

    rows = []
    for sentence in sentences:
        if passes_filters(sentence, target_pattern, excludes, min_len):
            rows.append(
                build_row(
                    document_name=pdf_path.name,
                    sentence=sentence,
                    lang=lang,
                    source_type="pdf",
                    extraction_method=extraction_method,
                )
            )

    return rows


# ---------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Extract candidate target sentences from Chinese or English policy "
            "documents in HTML or PDF format, optionally using OCR for scanned PDFs."
        )
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Input folder containing .pdf, .htm, or .html files",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output file path (.xlsx or .csv)",
    )
    parser.add_argument(
        "--source",
        choices=["pdf", "html"],
        required=True,
        help="Source type to process: pdf or html",
    )
    parser.add_argument(
        "--lang",
        choices=["cn", "en"],
        required=True,
        help="Document language: cn or en",
    )

    # OCR options (PDF only)
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Enable OCR fallback for scanned PDFs",
    )
    parser.add_argument(
        "--tesseract-lang",
        default="chi_sim",
        help="Tesseract language code for OCR (default: chi_sim)",
    )

    # Filtering and file search options
    parser.add_argument(
        "--min-len",
        type=int,
        default=10,
        help="Minimum sentence length (default: 10)",
    )
    parser.add_argument(
        "--max-cell-len",
        type=int,
        default=32767,
        help="Maximum cell length for Excel output (default: 32767)",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search the input folder recursively",
    )

    return parser.parse_args()


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------

def main() -> None:
    args = parse_args()

    input_dir = Path(args.input)
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    target_pattern, excludes, splitter = get_rules(args.lang)

    files = list(iter_files(input_dir, args.source, args.recursive))
    if not files:
        print(f"No matching {args.source.upper()} files found in: {input_dir}")
        return

    print(f"Found {len(files)} file(s) to process.")

    rows: List[Dict[str, str]] = []
    for path in files:
        if args.source == "html":
            rows.extend(
                extract_from_html_file(
                    html_path=path,
                    target_pattern=target_pattern,
                    excludes=excludes,
                    splitter=splitter,
                    min_len=args.min_len,
                    lang=args.lang,
                )
            )
        else:
            rows.extend(
                extract_from_pdf_file(
                    pdf_path=path,
                    target_pattern=target_pattern,
                    excludes=excludes,
                    splitter=splitter,
                    min_len=args.min_len,
                    enable_ocr=args.ocr,
                    tesseract_lang=args.tesseract_lang,
                    lang=args.lang,
                )
            )

    df = pd.DataFrame(rows)

    if df.empty:
        print("No candidate sentences found (0 rows).")
        return

    # Protect against Excel cell size limits.
    df = df[df["Sentence"].str.len() <= args.max_cell_len].copy()

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if out_path.suffix.lower() == ".xlsx":
        df.to_excel(out_path, index=False)
    elif out_path.suffix.lower() == ".csv":
        df.to_csv(out_path, index=False, encoding="utf-8")
    else:
        raise ValueError("Output file must end with .xlsx or .csv")

    print(f"Saved output to: {out_path}")
    print(f"Rows written: {len(df)}")


if __name__ == "__main__":
    main()