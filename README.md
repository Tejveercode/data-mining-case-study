# Data Mining Case Study — Hospital Patient Records

Course assignment for **Data Mining (Chandigarh University)**, session 2026–27.
Topic 3 of the case-study sheet: *a hospital wants to analyze patient records to
identify common patterns — explain how classification or clustering can help.*

The deliverable is `reports/Case_Study_Data_Mining.docx`, an editable Word file
(13 A4 pages) that answers that question directly, using real numbers computed
from a public heart-disease dataset.

## Before you print

Open the `.docx` in Word and fill in the four placeholders on the cover page:

- `[Your Name]`
- `[Your Roll Number]`
- `[Your Section]`
- `[Teacher's Name]`

Then print (or export to PDF). Everything else — page numbers, running header,
table of contents — is already in place. Word may refresh the table of contents
on open; that is expected and harmless.

## What is in the report

| Part | Content |
|---|---|
| Sections 1–3 | Introduction, the UCI Cleveland dataset (303 patients, 14 attributes), cleaning and encoding steps |
| Section 4 | Why clustering (K-means) and why classification (decision tree), plus reproduction steps |
| Sections 5–6 | Results: 2 patient groups (silhouette 0.237, reported honestly as modest) and a depth-4 decision tree (76.9% test accuracy vs. 53.9% baseline) |
| Section 7 | Two worked examples with real arithmetic: assigning a patient to a group, and walking the tree for a test patient |
| Section 8 | The direct answer to the assignment prompt: how both methods help the hospital |
| Sections 9–11 | Limitations, conclusion, references |

Figures: elbow/silhouette plot, PCA cluster scatter, confusion matrix, and the
full learned tree on a landscape page.

## Reproduce the analysis

```bash
python3 -m venv .venv
.venv/bin/pip install pandas scikit-learn matplotlib python-docx pypdf
.venv/bin/python src/analyze.py     # writes reports/results.json + 4 figures
.venv/bin/python src/build_report.py --toc-pages reports/toc_pages.json
```

`src/analyze.py` downloads the UCI Cleveland processed dataset, imputes the 6
missing values, one-hot encodes the categoricals, runs K-means (k chosen by
silhouette) and a decision tree (70/30 split, 10-fold CV), and saves every number
the report quotes to `reports/results.json`.

`src/build_report.py` generates the Word document from `results.json` alone —
change a number in the analysis and the report text follows. It needs
`python-docx`; the `--toc-pages` argument feeds cached page numbers into the
table of contents (regenerate it after any pagination change by extracting page
text from a PDF render and matching the heading titles).

## Layout rules used

- Booktabs-style tables: horizontal rules only, header row repeats across pages.
- Justified body text with hyphenation; headings numbered to two levels.
- Captions under every table and figure; figures generated at print resolution.
- Running header with title, footer page numbers (cover page suppressed).
- Real Word TOC field, so it can be refreshed in Word with one click.

## Repository layout

```
data/raw/processed.cleveland.data   UCI dataset (303 records)
src/analyze.py                      analysis pipeline
src/build_report.py                 report generator (python-docx)
reports/results.json                every number quoted in the report
reports/figures/*.png               the four figures
reports/Case_Study_Data_Mining.docx the deliverable
```

## References

1. Detrano, R. et al. (1989). *International application of a new probability
   algorithm for the diagnosis of coronary artery disease.* American Journal of
   Cardiology, 64(5), 304–310.
2. UCI Machine Learning Repository. *Heart Disease Data Set.*
   archive.ics.uci.edu/dataset/45/heart+disease
3. Kaggle mirror of the same dataset (1,025-row expanded version — not used here).
4. Hartigan, J. A., & Wong, M. A. (1979). *Algorithm AS 136: A K-Means
   clustering algorithm.* Journal of the Royal Statistical Society C, 28(1), 100–108.
5. Breiman, L. et al. (1984). *Classification and Regression Trees.* Wadsworth.
6. Pedregosa, F. et al. (2011). *Scikit-learn: Machine Learning in Python.*
   JMLR, 12, 2825–2830.
7. Kaufman, L., & Rousseeuw, P. J. (1990). *Finding Groups in Data.* Wiley.
   (silhouette coefficient)
8. Wolpert, D. H., & Macready, W. G. (1997). *No Free Lunch Theorems for
   Optimization.* IEEE TEC, 1(1), 67–82.
