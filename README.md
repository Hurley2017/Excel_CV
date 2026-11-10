# Excel_CV

Traditional CV reimagined in Microsoft Excel — built from the portfolio and résumé data in `../tushermondalcv`.

## Output

- **[Tusher_Mondal_CV.xlsx](Tusher_Mondal_CV.xlsx)** — 4-tab Excel workbook:
  1. **`CV`** — Traditional 2-page portrait CV layout crafted directly in the Excel grid (clickable contact/repo links, formula-driven snapshot bar, crisp typography, and print-ready page break).
  2. **`Career Dashboard`** — Management-reporting style scorecard with live cross-sheet formulas (`COUNTA`, `COUNTIF`, `AVERAGEIF`, `SUM`, `TODAY`, `ROUND`) and two native Excel charts.
  3. **`Experience & Projects`** — Career timeline with dynamic tenure formulas (`Months` computed against `TODAY()` for the current role) and a filterable 9-project ledger.
  4. **`Education & Skills`** — Academic record, competitive programming metrics, certifications, and a filterable 50-skill inventory across 6 domains.

## Regenerating

Requires Python 3.12+ with `openpyxl` and `pillow`:

```bash
python generate_cv.py
```
