# Research De-identification Workspace (RDW) v0.1

An early **Windows prototype** for de-identifying Word documents using an editable CSV dictionary. This is not a verified anonymization solution; all output needs human review.

## Setup

1. Install Python 3 with Tkinter support (standard Windows python.org installer).
2. Unzip this package into your GitHub repository folder (keep your existing files; merge folders).
3. Open Command Prompt in that folder.
4. Run `py -m pip install -r requirements.txt` (or `python -m pip install -r requirements.txt`).
5. Run `py run_rdw.py` (or `python run_rdw.py`).
6. Choose an input `.docx`, a CSV dictionary, and an output folder.

## Results

The output folder contains a de-identified `.docx`, an audit `.csv`, and a processing manifest `.json` with input and dictionary hashes.

## Important limitations

- v0.1 handles only DOCX; PDF, OCR, spreadsheets and batch processing are future milestones.
- Modified paragraphs can lose run-level formatting (bold, italics, hyperlinks).
- Text in drawings, text boxes, images, tracked changes, comments, embedded objects, and metadata is **not reliably de-identified**.
- URLs in hyperlink relationships may persist even when visible text is replaced. Do not publish the output before checking these.
- Names and abbreviations can have ambiguous meanings; replacement is case-insensitive and needs manual review.
- Audit logs contain **original identifying text** and must be stored securely and kept out of public GitHub repositories.
- Do not commit research documents, identifying dictionaries, or audit logs to a public repository.

## Next milestones

Preserve formatting, strip embedded links and metadata, add structured QC, then PDF and OCR support.
