import xlwings as xw 
import pandas as pd  
path = "/Users/speedy/Documents/GitProjects/AmpSheetRepository/AmpSheetMain/test2.xlsx" 
wbTest = xw.Book("test2.xlsx")

ws1 = wbTest.sheets(0)
# print(wbTest.range('A1).options(numbers=int).value) PRINTS AN INTEGER USE FOR CHANNEL NUMBERS
racksSheet = pd.read_excel(path, 'Racks')
wattsSheet = pd.read_excel(path, 'Watts').to_dict()
actualWatts = []
rowWatts = 0 # to factor that rows total watts
MAXCHANS = 25 # make it one more so less than is easier plus have to slice 5 off the front
cCounter = 0

for counter in range(3, 9, 1): # should be the rows that i am getting the channels from 
    cCounter += 1
    channelNums = []
    
    
    for k, v in racksSheet.items():
        
        channelNums.append(racksSheet[k][cCounter])
    print(channelNums[5:]) # this is where the slice happens and is working 
