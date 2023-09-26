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
channelNums = []
startingPositions = [(3,9,1), (16,22,1), (28,34,1), (41,47,1), (54,60,1), (67,73,1), (80,86,1), (93,99,1)]
channelNums4 = []
cCounter = 0
allChannels = []
posX = startingPositions[1][0]
print(posX)

for k, v in racksSheet.items():
    pass
row1 = racksSheet.loc[4]
row2 = data.iloc[3]
row3 = racksSheet.iloc[4]
print(row1)
print('***************************************************')
for k, v in row3.items():
    print(v)
    allChannels.append(v)
    print('***************************************************')
print(f"all channels : {allChannels[5:]}")

    

    










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
