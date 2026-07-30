import mysql.connector as mysql
import json

with open("booking_logs/all-bookings.json", "r") as _o:
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

for i, ob in enumerate(obs):
    print(json.dumps(ob, indent = 4 ))
    booked_on, booked_at = ob['booked_at'].split(" ")
    fullname = f"{ob['firstname']} {ob['middle']} {ob['lastname']}"
    charges = ""
    for charg in ob['charges']:
        charges += f"[{charg['code']}] {charg['name']}, "
    charges = charges[:-2]

    charges = charges.replace("'", "")
    sql = f"INSERT INTO bookings (lastname, middle, firstname, fullname, sex, dob, booked_on, booked_at, charges) VALUES ('{ob['lastname']}', '{ob['middle']}', '{ob['firstname']}', '{fullname}', '{ob['sex']}', '{ob['dob']}', '{booked_on}', '{booked_at}', '{charges}')"
    print(sql)
    query(sql)
