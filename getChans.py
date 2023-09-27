import pandas as pd  

STARTING_POSITIONS = [1,2,3,4,5,6,14,15,16,17,18,19,26,27,28,29,30,31,39,40,41,42,43,44,52, 53, 54, 55, 56, 57,65,66,67,68,69,70,78,79,80,81,82,83,91,92,93,94,95,96]

class Channels:
    def __init__(self) -> None:
        self.allChannels = []
        self.channelNums = []
        self.path = "/Users/speedy/Documents/GitProjects/AmpSheetRepository/AmpSheetMain/test2.xlsx"
        self.racksSheet = pd.read_excel(self.path, 'Racks')
        
    def getChans(self, allChans):
        for position in STARTING_POSITIONS:
            self.channelNums = []   # resets channel nums back to blank for next line. 
            row1 = self.racksSheet.loc[position] # Getting the row to append to AllChannels This will get all 48 circuits

            for k, v in row1.items():
                
                self.channelNums.append(v)
                
            self.allChannels.append(self.channelNums[5:]) # slices off the unneeded stuff leaves just the channels
        
        return self.allChannels # returns to main the channels of the sheet
    
    