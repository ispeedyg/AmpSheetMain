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
MAXCHANS = 20

for k, v in racksSheet.items():
    print(k, v)
