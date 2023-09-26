import os

def clearScreen():
    os.system('clear')
clearScreen()

import xlwings as xw 
import pandas as pd  
path = "/Users/speedy/Documents/GitProjects/AmpSheetRepository/AmpSheetMain/test2.xlsx" 
wbTest = xw.Book("test2.xlsx")

ws1 = wbTest.sheets(0)
# print(wbTest.range('A1).options(numbers=int).value) PRINTS AN INTEGER USE FOR CHANNEL NUMBERS
racksSheet = pd.read_excel(path, 'Racks')
wattsSheet = pd.read_excel(path, 'Watts').to_dict()
startingPositions = [(3,9,1), (16,22,1), (28,34,1), (41,47,1), (54,60,1), (67,73,1), (80,86,1), (93,99,1)]
cCounter = 0
channelNumsSliced = []
allChannels = []
# for position in startingPositions:
    

for counter in range(3, 9, 1): # should be the rows that i am getting the channels from 
    cCounter += 1
    channelNums = []
       
    for k, v in racksSheet.items():
        
        channelNums.append(racksSheet[k][cCounter])
        # channelNumsSliced = channelNums[5:]
    allChannels.append(channelNums[5:]) # creating a large list for the entire sheet to work from later
    print(channelNums[5:]) # this is where the slice happens and is working 
print(f"allchannels = {allChannels}")
