"""
Script for Auto Optimization for Formal Recruitment


Created 4/8/2020

History:
        Tim Sun     4/8/20
            Created
"""
import xlrd
import xlsxwriter

spreadsheet_location = "PNM Schedule 1 Censored.xlsx"

# participatingHouses = ["Pike", "KSig", "KA", "Lambda", "SigNu", "Sigs", "BYX"]
participatingHouses = ["A", "B", "C", "D", "E", "F", "G"]

# House: [Total(Functional Priority), Goal, Party 1, ... , Party n]
ChapterDictionary = {}

# PNM Name: [List to Visit]
PNMDictionary = {}

PNMPriorityList = []
ChapterPriorityList = []

visitingDictionary = {}

firstCol = 0
lastCol = 0
numrows = 0
numcols = 0
numparties = 0


"""
Get the spreadsheet that needs to be modified

Will need to take some sort of user input eventually, maybe same with Col Numbers
"""
def getSpreadsheet():
    workbook = xlrd.open_workbook(spreadsheet_location)
    spreadsheet = workbook.sheet_by_index(0)

    global firstCol
    global lastCol
    global numparties
    global numcols
    global numrows
    firstCol = 2
    lastCol = 5
    numcols = spreadsheet.ncols
    numrows = spreadsheet.nrows
    numparties = numcols - 2

    return spreadsheet


def initChapterDict():
    values = [0 for i in range(2 + numparties)]
    ChapterDictionary.update({key: list(values) for key in participatingHouses})


def countByHouse(spreadsheet):
    for y in range(1, numrows):
        name = spreadsheet.cell_value(y, 0)
        group = spreadsheet.cell_value(y, 1)
        visitList = []
        for x in range(firstCol, lastCol + 1):
            house = spreadsheet.cell_value(y, x)
            if house != "":
                visitList.append(house)
                ChapterDictionary[house][0] += 1
                # print(name + ": " + house)
            PNMDictionary.update({name: (group, [len(visitList)], visitList)})

    for chapter in ChapterDictionary:
        ChapterDictionary[chapter][1] = ChapterDictionary[chapter][0] / numparties


def establishPNMPriority():
    values = [0 for i in range(numparties)]
    for pnm in PNMDictionary:
        name = pnm
        priority = PNMDictionary[pnm][1]
        PNMPriorityList.append((name, priority, values))
    lst = len(PNMPriorityList)
    for i in range(0, lst):
        for j in range(0, lst-i-1):
            if PNMPriorityList[j][1] > PNMPriorityList[j+1][1]:
                temp = PNMPriorityList[j]
                PNMPriorityList[j] = PNMPriorityList[j + 1]
                PNMPriorityList[j + 1] = temp
    print("PNMPriorityList")
    print(PNMPriorityList)


def establishHousePriority():
    for house in ChapterDictionary:
        chapter = house
        priority = ChapterDictionary[house][0]
        ChapterPriorityList.append((chapter, priority))

    numHouses = len(ChapterPriorityList)
    for i in range(0, numHouses):
        for j in range(0, numHouses-i-1):
            if ChapterPriorityList[j][1] > ChapterPriorityList[j+1][1]:
                temp = ChapterPriorityList[j]
                ChapterPriorityList[j] = ChapterPriorityList[j + 1]
                ChapterPriorityList[j + 1] = temp
    print("ChapterPriorityList")
    print(ChapterPriorityList)


def addParty(pnm, house, availability):
    # determine first available party slot
    name = pnm[0]
    # visitingDictionary.update({pnm: schedule})
    # Compare to house values

    return


def makeSchedule():
    for pnm in PNMPriorityList:
        schedule = ()
        name = pnm[0]
        visiting = PNMDictionary[name][2]
        availability = pnm[2]
        # Figure out how we want to hash
        for house in ChapterPriorityList:   # Already goes by chapter priority
            if house[0] in visiting:
                addParty(pnm, house[0], schedule)
        visitingDictionary.update({name: schedule})


def writeSchedule():
    return


def main():
    sheet = getSpreadsheet()
    initChapterDict()
    countByHouse(sheet)
    establishPNMPriority()
    establishHousePriority()
    makeSchedule()

    writeSchedule()

    print("ChapterDictionary")
    print(ChapterDictionary)

    print("PNMDictionary")
    print(PNMDictionary)


"""
Establisha  breakpoint for when breaks are allowed (greater than 1/2 of parties probably)

Before breakpoint, work from least groups to most 

After breakpoint, work from most groups to least, keeping track of priority via the counters 

still need some sort of remapping program


Problems Right now: Its physically impossible to create balanced groups while maintaining that everyone visits a house group 1
At the same time, if all the houses can already handle the capacity of the first group, there should be no problem with
having groups get smaller as the day goes on


"""


if __name__ == "__main__":
    main()
