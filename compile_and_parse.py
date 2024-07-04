#Imports
from asyncio.windows_events import NULL
from pprint import pprint
from openpyxl import Workbook
from datetime import datetime

## This is to determine the month so we don't have to do as much maintenance.
manualorautomatic = int(input("Would you like to input month/year manually (1)\nor have it done automatically for todays year/month? (2)\nor have it done automatically for yesterdays year/month (3): "))
manualorautomaticpicked = NULL
match manualorautomatic:
    case 1:
        print("You have picked to input month/year manually.")
        manualorautomaticpicked = True
    case 2:
        print("You have picked to have the month/year part automated for todays year/month")
        manualorautomaticpicked = False
        currentMonth = datetime.now().month
        currentYear = datetime.now().year
    case 3:
        print("You have picked to have the month/year part automated for yesterdays year/month")
        manualorautomaticpicked = False
        currentMonth = datetime.now().month - 1
        if currentMonth == 0:
            currentYear = datetime.now() - 1
            currentMonth = 12
        else:
            currentYear = datetime.now().year
    case _:
        print("Invalid input - Exitting program.")
        exit()

if(manualorautomaticpicked):
    currentMonth = int(input("Pick the month of the year (1 is January and 12 is December): "))
    currentYear = int(input("Pick the year desired: "))

themonth = "TGMC_"
lastmonth = "TGMC_"
match currentMonth:
    case 0:
        themonth += "December"
        lastmonth += "November"
    case 1:
        themonth += "January"
        lastmonth += "December"
    case 2:
        themonth += "February"
        lastmonth += "January"
    case 3:
        themonth += "March"
        lastmonth += "February"
    case 4:
        themonth += "April"
        lastmonth += "March"
    case 5:
        themonth += "May"
        lastmonth += "April"
    case 6:
        themonth += "June"
        lastmonth += "May"
    case 7:
        themonth += "July"
        lastmonth += "June"
    case 8:
        themonth += "August"
        lastmonth += "July"
    case 9:
        themonth += "September"
        lastmonth += "August"
    case 10:
        themonth += "October"
        lastmonth += "September"
    case 11:
        themonth += "November"
        lastmonth += "October"
    case 12:
        themonth += "December"
        lastmonth += "November"
    case _:
        print("Invalid month picked - Exitting program.")
        exit()
    
# i really don't wanna maintain the month/year thingy
file_name_txt = f"./{currentYear}/{themonth}.txt"
file_name_txt_results = f"./{currentYear}_results/{themonth}.txt"
file_name_txt_results_previous = f"./{currentYear}_results/{lastmonth}.txt"
file_name_xlsx = f"./{currentYear}/{themonth}.xlsx"
file_name_xlsx_results = f"./{currentYear}_results/{themonth}.xlsx"

## Compiling data stuff for spreadsheet
file_name_config = f"./config.txt"

groundmaps = []
groundmapdicts = []
shipmaps = []
shipmapdicts = []
gamemodes = []
gamemodedicts = []
rounds = []
rounddicts = []
AllLists = groundmaps, shipmaps, gamemodes, rounds
mode = groundmaps

configfile = file_name_config
read_config = open(configfile, encoding="utf-8")

for line in read_config:
    if "GroundMaps" in line:
        mode = groundmaps
    elif "ShipMaps" in line:
        mode = shipmaps
    elif "Gamemodes" in line:
        mode = gamemodes
    elif "Rounds" in line:
        mode = rounds
    else:
        mode += [line.strip("\n")]

def CreateDictsOf(list, list2):
    for i in list:
        a = {}
        a[i] = {
            "name": i,
            "count": 0,
            "percentage": 0,
            "increase_decrease": 0,
        }
        list2.append(a)

CreateDictsOf(groundmaps,groundmapdicts)
CreateDictsOf(shipmaps,shipmapdicts)
CreateDictsOf(gamemodes,gamemodedicts)
CreateDictsOf(rounds,rounddicts)
# total count of ships/groundmaps/gamemodes so we can automate percentages
groundmapstotalcount = 0
shipmapstotalcount = 0
gamemodestotalcount = 0
#The file
datafile = file_name_txt
read_data = open(datafile, encoding="utf-8")

# Looking for stuff, probably terribly inefficient but hey i got-
# - a good enough computer, so it doesn't matter, right?
for line in read_data:
    for i in groundmaps:
        if i in line:
            groundmapdicts[groundmaps.index(i)][i]["count"] += 1
            groundmapstotalcount += 1
    for i in shipmaps:
        if i in line:
            shipmapdicts[shipmaps.index(i)][i]["count"] += 1
            shipmapstotalcount += 1
    for i in gamemodes:
        if i in line:
            gamemodedicts[gamemodes.index(i)][i]["count"] += 1
            gamemodestotalcount += 1
    for i in rounds:
        if i in line:
            rounddicts[rounds.index(i)][i]["count"] += 1

## Spreadsheet stuff
#stuff to make things easier for us here
wb = Workbook()
ws = wb.active
def thingstats(list1,list2):
    for i in list2:
        v = list1[list2.index(i)][i]["increase_decrease"]
        w = list1[list2.index(i)][i]["percentage"]
        x = list1[list2.index(i)][i]["count"]
        y = list1[list2.index(i)][i]["name"]
        z = [y,x,w,v]
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
def whatisit(str1, str2, str3, str4):
    ws.append([str1,str2, str3, str4])
    wb.save(file_name_xlsx_results)
def percentagemathing(list1,list2, var):
    for i in list2:
        if (list1[list2.index(i)][i]["count"] < 1):
            list1[list2.index(i)][i]["percentage"] = "N/A"
        else:
            list1[list2.index(i)][i]["percentage"] = (list1[list2.index(i)][i]["count"] / var) * 100
            list1[list2.index(i)][i]["percentage"] = round(list1[list2.index(i)][i]["percentage"], 1)
            list1[list2.index(i)][i]["percentage"] = str(list1[list2.index(i)][i]["percentage"])
            list1[list2.index(i)][i]["percentage"] += "%"
## The spreadsheet code
#Delete & unformat anything we previously had there
f = open(file_name_xlsx_results, "w")
f = open(file_name_txt_results, "w")
#counting percentages
percentagemathing(groundmapdicts,groundmaps, groundmapstotalcount)
percentagemathing(shipmapdicts,shipmaps, shipmapstotalcount)
percentagemathing(gamemodedicts,gamemodes, gamemodestotalcount)
#New stuff
whatisit("Post-round groundside map choices", "Times picked", "% picked","% increase/decrease")
thingstats(groundmapdicts,groundmaps)
whitespace()
whatisit("Post-round shipside map choices", "Times picked", "% picked","% increase/decrease")
thingstats(shipmapdicts,shipmaps)
whitespace()
whatisit("Gamemodes", "Times played", "% played","% Increase/decrease")
thingstats(gamemodedicts,gamemodes)
whitespace()
whatisit("Rounds played:", " ", " ", " ")
thingstats(rounddicts,rounds)
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
