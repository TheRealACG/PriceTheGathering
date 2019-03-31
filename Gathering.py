"""
Andrew Greer
Price The Gathering

3/30/2019
"""

import os
import csv
import xlrd
import xlwt

from bs4 import BeautifulSoup
import requests

#global variables, rows and columns 0 indexed
CARD_NAME_COLUMN = 1;
CARD_EXPANSION_COLUMN = 2;
CARD_PRICE_COLUMN = 4;

def cardNameCheck(name):
    name = name.replace(" ", "+")
    name = name.replace(",", "")
    return name
#end of cardNameCheck
# path of excel photo
loc = (r"C:\Users\Andrew\Documents\MTG card collection.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print(sheet.nrows)

print(sheet.cell_value(5,1))

for row in range(sheet.nrows - 1):
    cardName = sheet.cell_value(row + 1,CARD_NAME_COLUMN)
    cardName = cardNameCheck(cardName)
    if cardName == "":
        continue
    expansion = sheet.cell_value(row + 1,CARD_EXPANSION_COLUMN)
    print(cardName)
    print(expansion)
    startingURL = 'https://www.mtggoldfish.com/price/' + expansion + '/' + cardName + '#paper'

#expansion = "Theros"
#cardName = "Purphoros, God of the Forge"
#cardName = cardNameCheck(cardName)

source = requests.get(startingURL).text
soup = BeautifulSoup(source, 'lxml')
prices = soup.find("div", {"class" : "price-box paper"})
paperPrice = prices.find("div", {"class" : "price-box-price"})
print(paperPrice)
"""
for each_price in frames:
    print(each_price)
"""