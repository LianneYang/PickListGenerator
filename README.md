Pick List Generator
===================

A single-file Streamlit app that turns a stock file + a demand file into a
formatted Excel pick list.

WHAT IT DOES
------------
- Auto-detects whether the demand file has duplicate materials.
  • No duplicates → extends the flat demand sheet (renamed to "WO") with
    Pick Location, Pick QTY, Cumulative, Batch.
  • Duplicates    → builds WO, Summary PN, Pick List, and DN sheets.
- Allocates stock FIFO by batch; within a batch, the first pick chooses the
  bin whose qty is closest to demand, and each following pick makes the
  running cumulative closest to demand.
- Flags shortfalls with a bold "Miss Stock" row and the missed qty in red.
- Adds a "Problems" sheet when any material is short.

INPUTS
------
- Stock file  (.xlsx): Material, Storage Bin, Batch, Available stock
- Demand file (.xlsx): flat sheet with Schedule line date, Discharge Location,
  Material, Description, Quantity (plus optional Unrestricted-Use Stock,
  Comments, Customer Material Number, Sched.agreemnt)

USAGE
-----
Web (no install):
    https://picklistgenerator.streamlit.app

Local:
    pip install -r requirements.txt
    streamlit run WebApp_PickListGenerator.py

Then upload both files and click "Generate Pick List".

NOTES
-----
- Files are processed in memory; nothing is stored on the server.
- Requires Python 3.9+.
