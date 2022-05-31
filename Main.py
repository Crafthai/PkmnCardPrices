# import dependencies
import requests
# pull in api keys from config file
from config import public_key, private_key
# import spreadsheet stuff
import xlwt
import xlrd
from xlwt import Workbook
from xlutils.copy import copy

# construct api request
response = requests.post(
    "https://api.tcgplayer.com/token",

    headers={
        "Content-Type": "application/json",
        "Accept": "application/json"},

    data=(f"grant_type=client_credentials"
          f"&client_id={public_key}&"
          f"client_secret={private_key}")
)

# RUN ONCE PER DAY =========================================================================
day = 1  # curr date 5/31/22
# bearer token
access = response.json()['access_token']
# pass the info thru the header!!!
headers = {"accept": "application/json",
           "Content-Type": "application/json",
           'User-Agent': 'YOUR_USER_AGENT',
           "Authorization": "bearer " + access}

# sealed product list

# the reading file
wbr = xlrd.open_workbook("sealeds.xls")
# the writing file
wbw = copy(wbr)
sWrite = wbw.get_sheet(0)
sRead = wbr.sheet_by_index(0)
# max index of the # of sealed products in the spreadsheet ============================================================
SKUIndex = 9
sealedIndex = 9
sealedSKUs = []
sealedNames = []
sealedUids = []

# loop to create lists that contain the SKUs and names of cards
for x in range(SKUIndex):
    sealedSKUs.append(sRead.cell_value(1, x + 1))
    sealedNames.append(sRead.cell_value(0, x + 1))

for x in range(sealedIndex):
    sealedUids.append(sRead.cell_value(2, x + 1))

# check to see if there are more card uids than card skus, and if true then add the skus to the spreadsheet
if sealedIndex > len(sealedSKUs):
    i = sealedIndex - (sealedIndex - len(sealedSKUs))
    for x in range(sealedIndex - i):
        url = "https://api.tcgplayer.com/catalog/products/" + sealedUids[i] + "/skus"
        response = requests.get(url, headers=headers).json()
        # for calculating market price of NM cards
        curr_SKU = str(response['results'][0]['skuId'])
        # write the sku to the spreadsheet
        sWrite.write(1, x + 1, curr_SKU)
        i = i + 1
    # this is the name of the file for daily price updates
    wbw.save("sealeds.xls")

# loop to rebuild the lists if the skus and indexes didnt match
if SKUIndex != sealedIndex:
    cardSKUs = []
    cardNames = []
    for x in range(sealedIndex - 1):
        cardSKUs.append(sRead.cell_value(1, x + 1))
        cardNames.append(sRead.cell_value(0, x + 1))

i = 0
for x in sealedSKUs:
    url = "https://api.tcgplayer.com/pricing/marketprices/" + x
    response = requests.get(url, headers=headers)
    print("current market price of " + sealedNames[i] + ": " + str(response.json()['results'][0]['price']))
    sWrite.write(day, i + 1, str(response.json()['results'][0]['price']))
    i += 1
# this is the name of the file for daily price updates
wbw.save("sealeds.xls")

# run separately for sealed vs singles for different spreadsheet
# =====================================================================================================================
# =====================================================================================================================

# the reading file
cardReader = xlrd.open_workbook("cards.xls")
# the writing file
cardWriter = copy(cardReader)
# used to copy the spreadsheet before writing over it
sWrite = cardWriter.get_sheet(0)
sRead = cardReader.sheet_by_index(0)
# max index of the # of cards in the spreadsheet ======================================================================
SKUIndex = 59
cardIndex = 59
cardSKUs = []
cardNames = []
cardUids = []

# loop to create lists that contain the SKUs and names of cards
for x in range(SKUIndex):
    cardSKUs.append(sRead.cell_value(1, x + 1))
    cardNames.append(sRead.cell_value(0, x + 1))

for x in range(cardIndex):
    cardUids.append(sRead.cell_value(2, x + 1))

# check to see if there are more card uids than card skus, and if true then add the skus to the spreadsheet
if cardIndex > len(cardSKUs):
    i = cardIndex - (cardIndex - len(cardSKUs))
    for x in range(cardIndex - i):
        url = "https://api.tcgplayer.com/catalog/products/" + cardUids[i] + "/skus"
        response = requests.get(url, headers=headers).json()
        # for calculating market price of NM cards
        curr_SKU = str(response['results'][0]['skuId'])
        # write the sku to the spreadsheet
        sWrite.write(1, i + 1, curr_SKU)
        i += 1
    # this is the name of the file for daily price updates
    cardWriter.save("cards.xls")

# loop to rebuild the lists if the skus and indexes didnt match
if SKUIndex != cardIndex:
    cardSKUs = []
    cardNames = []
    for x in range(cardIndex - 1):
        cardSKUs.append(sRead.cell_value(1, x + 1))
        cardNames.append(sRead.cell_value(0, x + 1))

i = 0
for x in cardSKUs:
    url = "https://api.tcgplayer.com/pricing/marketprices/" + x
    response = requests.get(url, headers=headers)
    print("current market price of " + cardNames[i] + ": " + str(response.json()['results'][0]['price']))
    sWrite.write(day, i + 1, str(response.json()['results'][0]['price']))
    i += 1
# this is the name of the file for daily price updates
cardWriter.save("cards.xls")
