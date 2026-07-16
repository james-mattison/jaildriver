# SLO booking-log crawler

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python jaildriver.py --overwrite --timeout 60 --save-html
```

Each date is written as `booking_logs/YYYY-MM-DD.json`. The file contains a
bare JSON list, with one object per booking and all charge rows grouped under
that booking.

Successful HTML pages are saved only when `--save-html` is supplied. Failed
attempts are always saved under `booking_logs/failed_html/`.
