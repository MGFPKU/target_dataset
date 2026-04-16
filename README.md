📄 Climate Target Candidate Extractor

A rule-based Python script to extract candidate sentences containing climate targets from Chinese and English policy documents (HTML and PDF).

The tool is designed for high-recall screening: it identifies sentences that are likely to contain quantitative or time-bound targets, which can then be reviewed manually or used for downstream processing.

🔍 Overview
This script processes collections of policy documents and:

extracts raw text from HTML and PDF files
optionally applies OCR for scanned PDFs
splits text into sentences (language-specific rules)
applies regex-based filters to identify candidate target sentences
exports results to CSV or Excel

The output is a structured table of candidate sentences for further analysis.

⚠️ Important note
This is a heuristic, rule-based extraction tool, not a full target identification system.

✔️ Designed for broad coverage (recall)
❗ May include false positives
❗ May miss some valid targets
👉 Output should be reviewed manually or refined further
📂 Supported inputs
Languages:
Chinese (--lang cn)
English (--lang en)
Formats:
HTML (.htm, .html)
PDF (.pdf)

⚙️ Installation
1. Clone the repository
git clone https://github.com/yourusername/climate-target-extractor.git
cd climate-target-extractor
2. Install dependencies
pip install -r requirements.txt
3. Optional: OCR setup (for scanned PDFs)

If using --ocr, install:

pip install pytesseract pdf2image

Additionally:

Install Tesseract OCR
Install Poppler (required by pdf2image)

🚀 Usage
Basic structure
python extract_target_candidates.py \
  --input <input_folder> \
  --output <output_file> \
  --source <pdf|html> \
  --lang <cn|en>
Examples
Chinese HTML
python extract_target_candidates.py \
  --lang cn --source html \
  --input ./data/html_cn \
  --output ./outputs/cn_html.xlsx
Chinese PDF with OCR
python extract_target_candidates.py \
  --lang cn --source pdf --ocr \
  --tesseract-lang chi_sim \
  --input ./data/pdf_cn \
  --output ./outputs/cn_pdf.xlsx
English PDF
python extract_target_candidates.py \
  --lang en --source pdf \
  --input ./data/pdf_en \
  --output ./outputs/en_pdf.csv

📊 Output format
The script produces a table with the following columns:

Column	Description
Document	Source file name
Sentence	Extracted candidate sentence
Language	Language of the document (cn or en)
Source_Type	Input type (html or pdf)
Extraction_Method	html_text, pdf_text, or ocr

🧠 Methodology (brief)
Candidate sentences are identified using:

Sentence segmentation
Chinese: punctuation-based splitting (。！？；)
English: punctuation-based splitting (. ! ?)
Target detection rules
Regex patterns capturing:
years (e.g. “by 2030”, “到2025年”)
comparative baselines (e.g. “compared to 2005”)
percentage changes (e.g. “reduce by 40%”)
Filtering criteria
minimum sentence length
presence of numerical values
exclusion of structural elements (tables, appendices, dates)

This approach prioritises coverage over precision.

⚠️ Limitations
HTML extraction captures all visible text, including navigation elements
PDF extraction quality depends on document structure
OCR results depend on scan quality and language model
Regex patterns are not exhaustive and may require adaptation for other corpora
No semantic validation of targets is performed

🧩 Potential extensions
Named entity recognition (NER) for metrics and sectors
Structured parsing (baseline year, target year, magnitude)
Improved HTML content filtering
Integration with NLP pipelines

📜 License

MIT License

👤 Author

Florian Wengel