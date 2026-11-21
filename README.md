# docx-fingerprint

Inspect and rewrite OOXML (.docx) metadata to eliminate cross-file fingerprints.

When multiple .docx files are produced on the same machine or by the same tool, they share telltale metadata: author names, application identifiers, editing session IDs (RSIDs), timestamps, and more. This tool helps you audit and rewrite those fields.

## Install

No dependencies beyond Python 3.7+ standard library.

```bash
git clone https://github.com/USER/docx-fingerprint.git
cd docx-fingerprint
chmod +x docx_fingerprint.py
```

## Usage

### Inspect

Show metadata of one or more files and check for overlapping fingerprints:

```bash
python docx_fingerprint.py inspect file1.docx file2.docx file3.docx
```

Output includes: author, application, timestamps, RSIDs, embedded authors in tracked changes, and cross-file overlap warnings.

### Patch

Rewrite metadata fields:

```bash
# Change author and application
python docx_fingerprint.py patch file.docx \
    --author "Zhang Wei" \
    --application "Microsoft Office Word" \
    --app-version "16.0000"

# Randomize editing session IDs and set timestamps
python docx_fingerprint.py patch file.docx \
    --randomize-rsids \
    --created "2026-11-18T14:22:00Z" \
    --modified "2026-11-21T03:15:00Z" \
    --total-time 142

# Strip comments, tracked changes authors, and custom properties
python docx_fingerprint.py patch file.docx \
    --strip-comments \
    --strip-custom

# Full differentiation
python docx_fingerprint.py patch file.docx \
    --author "LiuYang" \
    --application "Microsoft Office Word" \
    --app-version "16.0000" \
    --template "Normal.dotm" \
    --total-time 142 \
    --randomize-rsids \
    --strip-comments \
    --strip-custom \
    --description ""
```

### Typical workflow

```bash
# 1. Inspect all files for shared fingerprints
python docx_fingerprint.py inspect *.docx

# 2. Patch each file with distinct metadata
python docx_fingerprint.py patch a.docx --author "Alice" --application "Microsoft Office Word" --randomize-rsids
python docx_fingerprint.py patch b.docx --author "Bob"   --application "LibreOffice Writer"    --randomize-rsids
python docx_fingerprint.py patch c.docx --author "Carol" --application "WPS Office"            --randomize-rsids

# 3. Verify no overlaps remain
python docx_fingerprint.py inspect a.docx b.docx c.docx
```

## What gets checked / patched

| Field | Where | Inspect | Patch |
|---|---|---|---|
| creator | core.xml | ✓ | ✓ |
| lastModifiedBy | core.xml | ✓ | ✓ |
| created / modified | core.xml | ✓ | ✓ |
| revision | core.xml | ✓ | ✓ |
| title / description | core.xml | ✓ | ✓ |
| Application | app.xml | ✓ | ✓ |
| AppVersion | app.xml | ✓ | ✓ |
| Template | app.xml | ✓ | ✓ |
| Company | app.xml | ✓ | ✓ |
| TotalTime | app.xml | ✓ | ✓ |
| RSID (session IDs) | settings.xml, document.xml | ✓ | ✓ (randomize) |
| comments.xml / people.xml | word/ | ✓ | ✓ (strip) |
| custom.xml | docProps/ | ✓ | ✓ (strip) |
| Embedded w:author | document.xml | ✓ | — |

## When metadata changes back

Saving a .docx in Word, WPS, or LibreOffice overwrites `lastModifiedBy`, `modified`, `Application`, and `TotalTime`, and appends new RSIDs. Run the tool again after any edit-and-save cycle.

## License

MIT
