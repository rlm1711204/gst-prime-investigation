# GST-Prime Preliminary Investigation

<p align="right">Created by <b>Ram Lakhan Meena, INSPECTOR CBIC</b></p>

A Google Colab notebook that turns GST-Prime Analytica exports into a preliminary investigation workbook, filed by GSTIN in your own Google Drive.

**Open the latest version in Colab (bookmark this link):**

https://colab.research.google.com/github/rlm1711204/gst-prime-investigation/blob/main/GST_Prime_Preliminary_Investigation.ipynb

- `GST_Prime_Preliminary_Investigation.ipynb` is always the current version. The version log is in its first cell.
- `versions/` keeps every released version, named with its version number and date.
- `engine/` holds the same engine code that the notebook writes out at Steps 3–5, for reading on GitHub.
- **BO comparison reports** (*Tax liabilities and ITC comparison*, one per financial year, from the BO portal) can be added alongside the GST-Prime exports. They are compared month by month with every GST-Prime figure and taken as final for the months they cover (`engine/gpe_bo.py`).
- **E-way bill document-wise summary** (GST-Prime → Taxpayer Profile → E-Way Bill), saved to Excel in any shape — one sheet per date window, a table starting in any column, or the whole page pasted into one cell. Overlapping windows are counted once, pages are matched to the GSTIN, and the leads (bills of supply carrying tax, high values taxed at almost nothing, cess, rate before and after 22-09-2025, HSN in GSTR-1 against the e-way bill rate) go to sheets 19H–19M (`engine/gpe_ewb.py`). Step 8 can also build the e-way bill workbook alone.
- `rate_master/GST_RATE_MASTER.xlsx` is the published GST rate master (one row per entry of a rate notification, with its source). Step 5A of the notebook can fetch it; `engine/gpe_rates.py` reads it in any other notebook:

```python
import gpe_rates
M = gpe_rates.load_master('GST_RATE_MASTER.xlsx')
gpe_rates.lookup(M['rows'], '7204', 'Oct-2025')
```

This repository holds **code only**. No taxpayer data is stored here: uploads, extracted data and reports stay in the Google Drive of whoever runs the notebook.
