import os

def clearScreen():
    os.system('clear')
clearScreen()

import xlwings as xw 
import pandas as pd  
path = "/Users/speedy/Documents/GitProjects/AmpSheetRepository/AmpSheetMain/test2.xlsx" 
wbTest = xw.Book("test2.xlsx")
data = pd.read_csv("./nba.csv")

ws1 = wbTest.sheets(0)
# print(wbTest.range('A1).options(numbers=int).value) PRINTS AN INTEGER USE FOR CHANNEL NUMBERS
racksSheet = pd.read_excel(path, 'Racks')

wattsSheet = pd.read_excel(path, 'Watts').to_dict()

startingPositions = [1,2,3,4,5,6,14,15,16,17,18,19,26,27,28,29,30,31,39,40,41,42,43,44,52, 53, 54, 55, 56, 57,65,66,67,68,69,70,78,79,80,81,82,83,91,92,93,94,95,96]

cCounter = 0
allChannels = []
for position in startingPositions:
    channelNums = []
    
    row1 = racksSheet.loc[position]

    print('******************** Starting a new ************************')
    print('***************************************************')
    for k, v in row1.items():
        print(v)
        print('---------------------------------------------')
        channelNums.append(v)
        print(f"channelNums: {channelNums}")
    allChannels.append(channelNums[5:])
    print('***************************************************')
    print(f"all channels : {allChannels}")

    

    










#     for m in range(4, 9, 1):
#         channelNums4.append(racksSheet.loc[k, m])
#         fourthSection = channelNums4[:5]
        
#     allChannels.append(channelNums[5:]) # creating a large list for the entire sheet to work from later
#     print(channelNums[5:]) # this is where the slice happens and is working 
# print(f"allchannels = {allChannels}")         
            
            
#     #         channelNums.append(racksSheet[67][m])
#     #     allChannels.append(channelNums[5:]) # creating a large list for the entire sheet to work from later
#     #     print(channelNums[5:]) # this is where the slice happens and is working 
#     # print(f"all channels = {allChannels}")
