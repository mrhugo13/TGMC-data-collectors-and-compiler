#Imports
from asyncio.windows_events import NULL
from pprint import pprint
from openpyxl import Workbook
from datetime import datetime

## This is to determine the month so we don't have to do as much maintenance.

manualorautomatic = input("Would you like to input month/year manually (Y) or have it done automatically for todays year/month? (N): ")
manualorautomaticpicked = NULL
match manualorautomatic:
    case "Y":
        print("You have picked to input month/year manually.")
        manualorautomaticpicked = True
    case "N":
        print("You have picked to have the month/year part automated")
        manualorautomaticpicked = False
        currentMonth = datetime.now().month
        currentYear = datetime.now().year
    case _:
        print("Invalid input - Exitting program.")
        exit()

if(manualorautomaticpicked):
    currentMonth = int(input("Pick the month of the year (1 is January and 12 is December): "))
    currentYear = int(input("Pick the year desired: "))

themonth = "TGMC_"
match currentMonth:
    case 0:
        themonth += "December"
    case 1:
        themonth += "January"
    case 2:
        themonth += "February"
    case 3:
        themonth += "March"
    case 4:
        themonth += "April"
    case 5:
        themonth += "May"
    case 6:
        themonth += "June"
    case 7:
        themonth += "July"
    case 8:
        themonth += "August"
    case 9:
        themonth += "September"
    case 10:
        themonth += "October"
    case 11:
        themonth += "November"
    case 12:
        themonth += "December"
    case _:
        print("Invalid month picked - Exitting program.")
        exit()

# i really don't wanna maintain the month/year thingy
file_name_txt = f"./{currentYear}/{themonth}.txt"
file_name_txt_results = f"./{currentYear}_results/{themonth}.txt"
file_name_xlsx = f"./{currentYear}/{themonth}.xlsx"
file_name_xlsx_results = f"./{currentYear}_results/{themonth}.xlsx"

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
    "name": "Magmoor Digsite IV",
    "count": 0
}
prisonstation = {
    "name": "Prison Station",
    "count": 0
}
whiskeyoutpost = {
    "name": "Whiskey Outpost",
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
    "name": "Icarus Military Port",
    "count": 0
}
orion = {
    "name": "Orion Military Outpost",
    "count": 0
}
polarcolony = {
    "name": "Rocinante Polar Colony",
    "count": 0
}
researchoutpost = {
    "name": "Research Outpost",
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
    "name": "Pillar of Spring",
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
datafile = file_name_txt
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
        whiskeyoutpost["count"] += 1
    elif chigusa["name"] in line:
        chigusa["count"] += 1
    elif icycaves["name"] in line:
        icycaves["count"] += 1
    elif icarus["name"] in line:
        icarus["count"] += 1
    elif orion["name"] in line:
        orion["count"] += 1
    elif polarcolony["name"] in line: # NEW MONTH: Hey is this still merged/TM'd/open?
        polarcolony["count"] += 1
    elif researchoutpost["name"] in line:
        researchoutpost["count"] += 1
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
#stuff to make things easier for us here
wb = Workbook()
ws = wb.active
def thingstats(dict):
    x = dict["count"]
    y = dict["name"]
    z = [y,x]
    a = str(z)
    pprint(z)
    f = open(file_name_txt_results, "a")
    f.write(f"{a}\n")
    f = open(file_name_xlsx_results, "a")
    ws.append(z)
    wb.save(file_name_xlsx_results)
def whitespace():
    ws.append([" "," "])
    wb.save(file_name_xlsx_results)
def whatisit(str1, str2, str3):
    ws.append([str1,str2, str3])
    wb.save(file_name_xlsx_results)
## The spreadsheet code
#Delete & unformat anything we previously had there
f = open(file_name_xlsx_results, "w")
f = open(file_name_txt_results, "w")
#New stuff
whatisit("Post-round groundside map choices", "Times picked","% increase/decrease")
thingstats(bigred)
thingstats(icecolony)
thingstats(v624)
thingstats(magmoor)
thingstats(prisonstation)
thingstats(chigusa)
thingstats(icycaves)
thingstats(icarus)
thingstats(whiskeyoutpost)
thingstats(orion)
thingstats(polarcolony)
thingstats(researchoutpost)
whitespace()
whatisit("Post-round shipside map choices", "Times picked","% increase/decrease")
thingstats(theseus)
thingstats(minerva)
thingstats(sulaco)
thingstats(pillars)
whitespace()
whatisit("Gamemodes", "Times played","% Increase/decrease")
thingstats(crash)
thingstats(distress)
thingstats(civilwar)
thingstats(nuclearwar)
whitespace()
whatisit("Rounds played:", " ", " ")
thingstats(roundsamount)
## Style stuff for spreadsheet
# this just adjusts columns to fit our data
# stolen from velis @ https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size
dims = {}
for row in ws.rows:
    for cell in row:
        if cell.value:
            dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))    
for col, value in dims.items():
    ws.column_dimensions[col].width = value

# coloring stuff in to look prettier
# todo: i hate colors

# saving all the style stuff
wb.save(file_name_xlsx_results)
f.close() #just to make sure we have everything closed at the end of this program
