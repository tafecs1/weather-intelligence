Put `TEHSIL_MASTER.csv` here (export of the TEHSIL_MASTER tab: File > Download > CSV).
weather.py reads `data/masters/TEHSIL_MASTER.csv` first, so with this file present the
pipeline never opens any other project's Google Sheet.
