#!/bin/bash

HERE="$( dirname $( readlink -f "$0" ) )"

PYTHON="${HERE}/.venv/bin/python3"
echo Deleting todays entry, to refresh it
rm -f "${HERE}/booking_logs/$( date "+%Y-%m-%d").json"

echo "RUNNING JAILDRIVER"
$PYTHON jaildriver.py

echo "Inserting bookings into DB."
$PYTHON insert_today.py
