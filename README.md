# AE210 Final Project Web Grader (v2.0)

Browser grader for the AE210 Final Project Jet11 workbook. Mirrors the MATLAB final project autograder checks.

## Usage

1. Open `docs/index.html` (or serve the `docs/` folder).
2. Drop a Final Project Jet11 workbook using the provided course template.
3. Review score, deductions, and bonus items; everything stays client-side.

## Parity testing

Use `docs/test_runner.html` with expected outputs to compare against MATLAB logs.

## Notes

- Supporting documents/templates (Excel templates, RFP, etc.) live in `CommonAssets/` and are symlinked in.
- The `docs/` folder here is a local copy (no longer shared), so edits apply only to the Final Project web grader.
- Archival prep: verify symlinks to `CommonAssets/*` remain valid; keep this folder and `CommonAssets` together when moving/zip’ing; note `docs` is local.
