#!/usr/bin/env python3
import mysql.connector as mysql
import json
from datetime import datetime
import yaml

today = datetime.now().strftime("%Y-%m-%d")


with open(f"booking_logs/{today}.json", "r") as _o:
    obs = json.load(_o)


with open("/opt/jaildriver/config.yaml", "r") as _o:
    cnf = yaml.load(_o, Loader = yaml.Loader)


conn = mysql.Connect(host = cnf['db']['host'], 
                     user = cnf['db']['user'],
                     password = cnf['db']['password'],
                     database = cnf['db']['database'],
                     autocommit = True)
curs = conn.cursor(dictionary = True, buffered = True)

def query(sql, results = False):
    curs.execute(sql)
    if results:
        return curs.fetchall()
duplicates = 0
new = 0

i = 0
for i, ob in enumerate(obs):
    #print(json.dumps(ob, indent = 4 ))
    booked_on, booked_at = ob['booked_at'].split(" ")
    fullname = f"{ob['firstname']} {ob['middle']} {ob['lastname']}"
    charges = ""
    for charg in ob['charges']:
        charges += f"[{charg['code']}] {charg['name']}, "
    charges = charges[:-2]

    charges = charges.replace("'", "")
    exists = f"SELECT * FROM bookings WHERE fullname = '{fullname}' AND booked_at = '{booked_at}' AND booked_on = '{booked_on}'"
    already = query(exists, True)
    if already:
        duplicates += 1
        print(f"Duplicate: {ob['lastname']} {ob['firstname']} {ob['middle']} {booked_at} {booked_on}")
        continue
    else:
        sql = f"INSERT INTO bookings (lastname, middle, firstname, fullname, sex, dob, booked_on, booked_at, charges) VALUES ('{ob['lastname']}', '{ob['middle']}', '{ob['firstname']}', '{fullname}', '{ob['sex']}', '{ob['dob']}', '{booked_on}', '{booked_at}', '{charges}')"
        print(sql)
        query(sql)
        new += 1


print(f"Inserted {new} new, {duplicates} duplicates")
