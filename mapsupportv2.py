import re

# Array of map tile letters and corresponding offsets for eastings and northings
mapTiles = [["NY",3,5],
            ["NU",4,6],
            ["NT",3,6],
            ["NZ",4,5],
            ["NS",2,6],
            ["NX",2,5],
            ["SC",2,4],
            ["SD",3,4],
            ["SE",4,4],
            ["TA",5,4],
            ["OV",5,5]]

# Functions to translate grid reference into eastings and northings values in metres
#
def getEastingsAndNorthings (gridRef):
    "This returns the eastings element of a grid reference"
    offsets = translateGridLetters(gridRef)
    coordsInSquare = getCoords(gridRef)
    if offsets[0] == -1 or coordsInSquare[0] == -1:
        return [-1,-1]
    eastings = (offsets[0] * 100000) + coordsInSquare [0]
    northings = (offsets[1] * 100000) + coordsInSquare [1]
    return [eastings, northings]

def translateGridLetters (gridRef):
    "This returns the offset eastings and northings for a set of grid letters"
    letters = gridRef[0:2].upper()
    for mapTile in mapTiles:
        mapTileLetters = mapTile[0]
        if letters == mapTileLetters:
            return [mapTile[1],mapTile[2]]

    print "Don't know about letters " + letters
    return [-1,-1]

def getCoords (gridRef):
    "This splits grid reference number part into eastings and northings at 1m resolution"
    numbers = gridRef[2:]
    if len(numbers)%2 <> 0:
        print "Error - odd number of grid reference numbers"
        return [-1,-1]
    else:
        halfLength = len(numbers)/2
        eElement = int(numbers[0:halfLength])
        nElement = int(numbers[halfLength:])
        power = 5 - halfLength
        multiplier = 10 ** power
        eElement *= multiplier
        nElement *= multiplier
        return [eElement, nElement]

def getLWSMappings ():
    "This reads LWS mappings from file and creates a dictionary keyed on site description and id"
    mappingsFile = open(r"\\parsons\data\ERIC North East\IT\ArcMap Python stuff\LWS file mappings 3 - for script.txt","r")
    mappings = {}
    for line in mappingsFile:
        parts = line.split("\t")
        sitename = parts[0].replace('"','')    # 23/03/2016 added .replace('"','')
        siteid = parts[1]
        sitefile = parts[3].strip().replace('"','')
        key = sitename + siteid
        mappings[key] = sitefile
    return mappings

def validateGridRef (gridRef):
    "This validates a grid reference as a valid format with at least 6 digits"
    p1 = re.compile('[a-zA-Z]{2}([0-9]*)$')

    m = p1.match(gridRef)
    if not m:
        return False
    else:
        nums = m.group(1)
        if len(nums) % 2 <> 0 or len(nums) < 6:
            return False
        else:
            return True
