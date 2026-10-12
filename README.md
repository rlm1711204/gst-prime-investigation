# GST-Prime Preliminary Investigation

<p align="right">Created by <b>Ram Lakhan Meena, INSPECTOR CBIC</b></p>

A Google Colab notebook that turns GST-Prime Analytica exports into a preliminary investigation workbook, filed by GSTIN in your own Google Drive.

**Open the latest version in Colab (bookmark this link):**

https://colab.research.google.com/github/rlm1711204/gst-prime-investigation/blob/main/GST_Prime_Preliminary_Investigation.ipynb

- `GST_Prime_Preliminary_Investigation.ipynb` is always the current version. The version log is in its first cell.
- `versions/` keeps every released version, named with its version number and date.
- `engine/` holds the same engine code that the notebook writes out at Steps 3–5, for reading on GitHub.
- `rate_master/GST_RATE_MASTER.xlsx` is the published GST rate master (one row per entry of a rate notification, with its source). Step 5A of the notebook can fetch it; `engine/gpe_rates.py` reads it in any other notebook:

```python
import gpe_rates
M = gpe_rates.load_master('GST_RATE_MASTER.xlsx')
gpe_rates.lookup(M['rows'], '7204', 'Oct-2025')
```

This repository holds **code only**. No taxpayer data is stored here: uploads, extracted data and reports stay in the Google Drive of whoever runs the notebook.
