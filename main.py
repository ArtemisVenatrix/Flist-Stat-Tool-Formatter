import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

sheetName = input("Sheet Name: ")
knownColors = [
    "red",
    "purple",
    "blue",
    "black",
    "white",
    "pink",
    "brown",
    "yellow",
    "green"
]

def getSheet():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    sheet = client.open(sheetName).sheet1

    return sheet

def getCoreStats(sheet):
    names = sheet.range('A2:A11')
    finalPoints = sheet.range('C2:C11')
    modifiers = sheet.range('D2:D11')
    compiledList = {}

    for i, x in enumerate(names):
        compiledList[x.value] = (finalPoints[i].value, modifiers[i].value)

    return compiledList

def getSkills(sheet):
    names = sheet.range('A14:A26')
    proficiencies = sheet.range('C14:C26')
    finalPoints = sheet.range('D14:D26')
    compiledList = {}

    for i, x in enumerate(names):
        compiledList[x.value] = (proficiencies[i].value != "", finalPoints[i].value)

    return compiledList

def getMiscStats(sheet):
    names = sheet.range('F14:F15')
    stats = sheet.range('G14:G15')
    compiledList = {}

    for i, x, in enumerate(names):
        compiledList[x.value] = stats[i].value

    return compiledList

def requestColors():
    bodyColor = requestBodyColor()
    headerColor = requestHeaderColor()
    return bodyColor, headerColor

def requestBodyColor():
    color = input("What color would you like your stat main text to be? ").lower()
    if color not in knownColors:
        print("Unrecognised color, please try again.")
        requestBodyColor()
    return color

def requestHeaderColor():
    color = input("What color would you like your stat header text to be? ").lower()
    if color not in knownColors:
        print("Unrecognised color, please try again")
        requestHeaderColor()
    return color

def output(coreStats, skills, miscStats, sheetName, bodyColor, headerColor):
    with open(sheetName + ".txt", "w") as f:
        toWrite = "[center][u]Stats[/u][/center]\n[left]"
        toWrite += "[u]Stats[/u]" + " "*15 + "[u]Score[/u]" + " "*15 + "[u]Modifier[/u]\n"
        for x in coreStats:
            toWrite += "[u]" + x + ":[/u]"
            toWrite += " "*(20-len(x))
            toWrite += coreStats[x][0]
            toWrite += " "*(20-len(coreStats[x][0]))
            toWrite += coreStats[x][1]
            toWrite += " "*(20-len(coreStats[x][1]))
            toWrite += "\n"

        toWrite += "[/left]\n[center][u]Skills[/u][/center]\n[left]"
        toWrite += "[u]Skills[/u]" + " "*14 + "[u]Proficiency[/u]" + " "*9 + "[u]Modifier[/u]\n"

        for x in skills:
            toWrite += "[u]" + x + ":[/u]"
            toWrite += " "*(20-len(x))
            if skills[x][0]:
                toWrite += "yes"
                toWrite += " "*17
            else:
                toWrite += "no"
                toWrite += " "*18
            toWrite += skills[x][1]
            toWrite += " "*(20-len(skills[x][1]))
            toWrite += "\n"

        toWrite += "[/left]\n[center][u]Misc Stats[/u][/center]\n[left]"
        toWrite += "[u]Stats[/u]" + " "*15 + "[u]Score[/u]\n"

        for x in miscStats:
            toWrite += "[u]" + x + ":[/u]"
            toWrite += " "*(20-len(x))
            toWrite += miscStats[x]
            toWrite += "\n"

        toWrite += "[/left]"

        f.write(toWrite)

def main():
    sheet = getSheet()
    coreStats = getCoreStats(sheet)
    skills = getSkills(sheet)
    miscStats = getMiscStats(sheet)
    colors = requestColors()
    pprint(coreStats)
    print()
    pprint(skills)
    print()
    pprint(miscStats)
    output(coreStats, skills, miscStats, sheetName.replace(" ", "_"), colors[0], colors[1])

main()