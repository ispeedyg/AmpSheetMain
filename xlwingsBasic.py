
import pandas as pd  
path = "/Users/speedy/Documents/GitProjects/AmpSheetRepository/AmpSheetMain/test2.xlsx" 

# print(wbTest.range('A1).options(numbers=int).value) PRINTS AN INTEGER USE FOR CHANNEL NUMBERS
racksSheet = pd.read_excel(path, 'Racks')
# wattsSheet = pd.read_excel(path, 'Watts').to_dict()
# Below gets all the channels even if there is more than 5
startingPositions = [1,2,3,4,5,6,14,15,16,17,18,19,26,27,28,29,30,31,39,40,41,42,43,44,52, 53, 54, 55, 56, 57,65,66,67,68,69,70,78,79,80,81,82,83,91,92,93,94,95,96]

allChannels = []
for position in startingPositions:
    channelNums = []   # resets channel nums back to blank for next line. 
    row1 = racksSheet.loc[position + 0] # TODO maybe try a range loop plus (+) the range to the position so for x in range(1,7,1) position + x then reset x

    # print('******************** Starting a new ************************')
    # print('***************************************************')
    for k, v in row1.items():
        # print(v)
        # print('---------------------------------------------')
        channelNums.append(v)
        # print(f"channelNums: {channelNums}")
    allChannels.append(channelNums[5:]) # slices off the unneeded stuff leaves just the channels
print('***************************************************')
print(f"all channels : {allChannels}")


