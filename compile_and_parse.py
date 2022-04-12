######################################################################################
###                                                                                ###
# Quick note: all the [YEAR] and [MONTH] stuff isn't actual code, it's just meant to #
# be the place where you input the current month and year.                           #
###                                                                                ###
######################################################################################
#Imports
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

## Compiling data stuff for spreadsheet

# What we're looking for and where + how many maps we've found
# Groundside maps
icecount = 0
bigred_count = 0
lv624_count = 0
magmoor_count = 0
prison_count = 0
whisk_count = 0
chigusa_count = 0
icycaves_count = 0
icarus_count = 0
orion_count = 0
# Ship maps
theseus_count = 0
minerva_count = 0
sulaco_count = 0
pillars_count = 0
# Gamemode
crash_count = 0
distress_count = 0
civwar_count = 0
nuclearwar_count = 0
# Amount of rounds
roundscount = 0
#The file
datafile = "./[YEAR]/TGMC_[MONTH].txt" # NEW MONTH: Change to appropriate [YEAR]/TGMC_[MONTH].txt file
read_data = open(datafile, encoding="utf-8")

# Looking for stuff, probably terribly inefficient but hey you got-
# - a good enough computer, so it doesn't matter, right?
for line in read_data:
    if "Ice Colony" in line:
        icecount += 1
    elif "Big Red" in line:
        bigred_count += 1
    elif "LV624" in line:
        lv624_count += 1
    elif "Magmoor" in line:
        magmoor_count += 1
    elif "Prison Station" in line:
        prison_count += 1
    elif "Whiskey Outpost" in line:
        whisk_count += 1
    elif "Chigusa" in line:
        chigusa_count += 1
    elif "Icy Caves" in line:
        icycaves_count += 1
    elif "Icarus" in line:
        icarus_count += 1
    elif "Orion Military Outpost" in line: # NEW MONTH: Hey is this still merged/TM'd/open?
        orion_count += 1
    if "Theseus" in line:
        theseus_count += 1
    elif "Minerva" in line:
        minerva_count += 1
    elif "Sulaco" in line:
        sulaco_count += 1
    elif "Pillar" in line:
        pillars_count += 1
    if "Crash" in line:
        crash_count += 1
    elif "Distress" in line:
        distress_count += 1
    elif "Nuclear War" in line:
        nuclearwar_count += 1
    elif "Civil War" in line:
        civwar_count += 1
    if "Round ID" in line:
        roundscount += 1

## Spreadsheet stuff

#Things we need for spreadsheet to run
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) #hey you remembered to watch the video in the README.md yeah?
client = gspread.authorize(creds)
# Format i personally use is "TGMC_Datasheet_[Month]_[Year]"
sheet = client.open("TGMC_Datasheet_[Month]_[Year]").sheet1 # NEW MONTH: Change to appropriate month/year sheet 
#List of groundside, shipside maps, crash/distress rounds, and total amount of rounds + the amount we found
BigRed = ["Big Red", bigred_count]
IceCo = ["Ice Colony", icecount]
V624 = ["LV624", lv624_count]
Magmoor = ["Magmoor", magmoor_count]
PrisonStation = ["Prison Stat.", prison_count]
Whiskey = ["Whiskey-Post.", whisk_count]
Chigusa = ["Chigusa", chigusa_count]
IcyCaves = ["Icy Caves", icycaves_count]
Icarus = ["Icarus", icarus_count]
Orion = ["Orion M.O.", orion_count]
Theseus = ["Theseus", theseus_count]
Minerva = ["Minerva", minerva_count]
Sulaco = ["Sulaco", sulaco_count]
Pillars = ["PoS.", pillars_count]
Crash = ["Crash", crash_count]
Distress = ["Distress", distress_count]
Civilwar = ["Civil War", civwar_count]
Nuclearwar = ["Nuclear War",nuclearwar_count]
Totalrounds = ["Total Rounds:", roundscount]
quicksleep = time.sleep(2) # just to not bother with typing time.sleep(2) everytime, makes sure google isn't angry at us for writing too fast.
## The spreadsheet code
#Delete & unformat anything we previously had there
sheet.delete_rows(1,24)
#New stuff
###
# by the way, % +/- difference is done manually from last month's pie charts, this code doesn't
# actually do anything related to % +/- differences.
###
sheet.insert_row(['Post-round groundside map picks:'])
sheet.update_cell(1,3, '% +/- difference')
sheet.insert_row(BigRed,2)
sheet.insert_row(IceCo,3)
quicksleep
sheet.insert_row(V624,4)
sheet.insert_row(Magmoor,5)
sheet.insert_row(PrisonStation,6)
sheet.insert_row(Chigusa,7)
quicksleep
sheet.insert_row(IcyCaves,8)
sheet.insert_row(Icarus,9)
sheet.insert_row(Whiskey,10)
sheet.insert_row(Orion,11)
quicksleep
sheet.insert_row(['Post-round shipside map picks:'],13)
sheet.update_cell(13,3, '% +/- difference')
sheet.insert_row(Theseus,14)
sheet.insert_row(Minerva,15)
quicksleep
sheet.insert_row(Sulaco,16)
sheet.insert_row(Pillars,17)
sheet.insert_row(['Round gamemodes:'],19)
sheet.update_cell(19,3, '% +/- difference')
quicksleep
sheet.insert_row(Crash,20)
sheet.insert_row(Distress,21)
sheet.insert_row(Civilwar,22)
sheet.insert_row(Nuclearwar,23)
quicksleep
sheet.insert_row(Totalrounds,24)
sheet.update_cell(10,7,'All data here was collected between:')
sheet.update_cell(11,7,'[MONTH] 1st and [MONTH] [END OF MONTH DAY], year [YEAR].') # NEW MONTH: Update month (& Year if needed)
sheet.format(
    "E9:I12", 
    {"backgroundColor": {"red": 1.0,"green": 0.0,"blue": 0.0},
    "horizontalAlignment": "CENTER",
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0},"fontSize": 12,"bold": True}})
quicksleep
sheet.format(
    "A2:A11",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
sheet.format(
    "A13:A17",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
sheet.format(
    "A19:A24",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
sheet.format(
    "A1:C1",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
quicksleep
sheet.format(
    "A13:C13",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
sheet.format(
    "A19:C19",
    {"backgroundColor": {"red": 0.380,"green": 0.380,"blue": 0.380},
    "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 0.0}}})
sheet.format("B2:C2", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B4:C4", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
quicksleep
sheet.format("B6:C6", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B8:C8", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B10:C10", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B14:C14", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
quicksleep
sheet.format("B16:C16", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B20:C20", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B22:C22", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
sheet.format("B24:C24", {"backgroundColor": {"red": 0.811,"green": 0.811,"blue": 0.811}})
