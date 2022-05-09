######################################################################################
###                                                                                ###
# Quick note: all the [YEAR] and [MONTH] stuff isn't actual code, it's just meant to #
# be the place where you input the current month and year.                           #
###                                                                                ###
######################################################################################
#Imports
from pprint import pprint
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

## Compiling data stuff for spreadsheet

# dicts of what we're looking for + how many matching names of those dicts we've found
# Groundside maps
icecolony = {
    "name": "Ice Colony",
    "count": 0
}
bigred = {
    "name": "Big Red",
    "count": 0
}
v624 = {
    "name": "LV624",
    "count": 0
}
magmoor = {
    "name": "Magmoor",
    "count": 0
}
prisonstation = {
    "name": "Prison Station",
    "count": 0
}
whiskeyoutpost = {
    "name": "Whiskey",
    "count": 0
}
chigusa = {
    "name": "Chigusa",
    "count": 0
}
icycaves = {
    "name": "Icy Caves",
    "count": 0
}
icarus = {
    "name": "Icarus",
    "count": 0
}
orion = {
    "name": "Orion",
    "count": 0
}
# Ship maps
theseus = {
    "name": "Theseus",
    "count": 0
}
minerva = {
    "name": "Minerva",
    "count": 0
}
sulaco = {
    "name": "Sulaco",
    "count": 0
}
pillars = {
    "name": "Pillars",
    "count": 0
}
# Gamemode
crash = {
    "name": "Crash",
    "count": 0
}
distress = {
    "name": "Distress",
    "count": 0
}
civilwar = {
    "name": "Civil War",
    "count": 0
}
nuclearwar = {
    "name": "Nuclear War",
    "count": 0
}
# Amount of rounds
roundsamount = {
    "name": "Round ID",
    "count": 0
}
#The file
datafile = "./[YEAR]/TGMC_[MONTH].txt" # NEW MONTH: Change to appropriate [YEAR]/TGMC_[MONTH].txt file
read_data = open(datafile, encoding="utf-8")

# Looking for stuff, probably terribly inefficient but hey i got-
# - a good enough computer, so it doesn't matter, right?
for line in read_data:
    if icecolony["name"] in line:
        icecolony["count"] += 1
    elif bigred["name"] in line:
        bigred["count"] += 1
    elif v624["name"] in line:
        v624["count"] += 1
    elif magmoor["name"] in line:
        magmoor["count"] += 1
    elif prisonstation["name"] in line:
        prisonstation["count"] += 1
    elif whiskeyoutpost["name"] in line:
        whiskeyoutpost["name"] += 1
    elif chigusa["name"] in line:
        chigusa["count"] += 1
    elif icycaves["name"] in line:
        prisonstation["count"] += 1
    elif icarus["name"] in line:
        icarus["count"] += 1
    elif orion["name"] in line: # NEW MONTH: Hey is this still merged/TM'd/open?
        orion["count"] += 1
    if theseus["name"] in line:
        theseus["count"] += 1
    elif minerva["name"] in line:
        minerva["count"] += 1
    elif sulaco["name"] in line:
        sulaco["count"] += 1
    elif pillars["name"] in line:
        pillars["count"] += 1
    if crash["name"] in line:
        crash["count"] += 1
    elif distress["name"] in line:
        distress["count"] += 1
    elif nuclearwar["name"] in line:
        nuclearwar["count"] += 1
    elif civilwar["name"] in line:
        civilwar["count"] += 1
    if roundsamount["name"] in line:
        roundsamount["count"] += 1

## Spreadsheet stuff

#Things we need for spreadsheet to run
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #hey you remembered to watch the video in the README.md yeah?
client = gspread.authorize(creds)
# Format i personally use is "TGMC_Datasheet_[Month]_[Year]"
sheet = client.open("TGMC_Datasheet_[Month]_[Year]").sheet1 # NEW MONTH: Change to appropriate month/year sheet 

#stuff to make things easier for us here
sleep_delay = 2 # less typing
def shade(str):
    time.sleep(sleep_delay)
    sheet.format(str, {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
def altshade(str):
    time.sleep(sleep_delay)
    sheet.format(str,{"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},"textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
def thingstats(dict,index):
    x = dict["count"]
    y = dict["name"]
    z = [y,x]
    a = str(z)
    pprint(z)
    f = open("[YEAR]_results/TGMC_[MONTH].txt", "a") #NEW MONTH: check year/month
    f.write(a + "\n")
    f.close
    time.sleep(sleep_delay)
    sheet.insert_row(z,index)
## The spreadsheet code
#Delete & unformat anything we previously had there
sheet.delete_rows(1,24)
f = open("[YEAR]_results/TGMC_[MONTH].txt", "w") #NEW MONTH: check year/month
f.close
#New stuff
###
# by the way, % +/- difference is done manually from last month's pie charts, this code doesn't
# actually do anything related to % +/- differences.
###
sheet.insert_row(['Post-round groundside map picks:'])
sheet.update_cell(1,3, '% +/- difference')
thingstats(bigred,2)
thingstats(icecolony,3)
quicksleep
thingstats(v624,4)
thingstats(magmoor,5)
thingstats(prisonstation,6)
thingstats(chigusa,7)
quicksleep
thingstats(icycaves,8)
thingstats(icarus,9)
thingstats(whiskeyoutpost,10)
thingstats(orion,11)
quicksleep
sheet.insert_row(['Post-round shipside map picks:'],13)
sheet.update_cell(13,3, '% +/- difference')
thingstats(theseus,14)
thingstats(minerva,15)
quicksleep
thingstats(sulaco,16)
thingstats(pillars,17)
sheet.insert_row(['Round gamemodes:'],19)
sheet.update_cell(19,3, '% +/- difference')
quicksleep
thingstats(crash,20)
thingstats(distress,21)
thingstats(civilwar,22)
thingstats(nuclearwar,23)
quicksleep
thingstats(roundsamount,24)
sheet.update_cell(10,7,'All data here was collected between:')
sheet.update_cell(10,7,'All data here was collected between:')
sheet.update_cell(11,7,'[MONTH] 1st and [MONTH] [END OF MONTH DAY], year [YEAR].') # NEW MONTH: Update month (& Year if needed)
sheet.format(
    "E9:I12", 
    {"backgroundColor": {"red": 1.0,"green": 0.0,"blue": 0.0},
    "horizontalAlignment": "CENTER",
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0},"fontSize": 12,"bold": True}})
quicksleep
altshade("A2:A11")
altshade("A13:A17")
altshade("A19:A24")
altshade("A1:C1")
quicksleep
altshade("A13:C13")
altshade("A19:C19")
shade("B2:C2")
shade("B4:C4")
quicksleep
shade("B6:C6")
shade("B8:C8")
shade("B10:C10")
shade("B14:C14")
quicksleep
shade("B16:C16")
shade("B20:C20")
shade("B22:C22")
shade("B24:C24")
