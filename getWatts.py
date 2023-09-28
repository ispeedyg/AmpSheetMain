import pandas as pd 

class Watts:
    def __init__(self) -> None:
        self.chanLow = int
        self.chanHi = int
        self.watts = int
        self.lineWatts = 0
        self.totalLineWatts = []
        self.counter = 0
        self.temp = []
        
     
    def whatsWatts(self, path, list):
        data = pd.read_excel(path, 'Watts').to_dict()
        for subList in range(len(list)):
            self.temp.append(list[subList])
    
        
        # print(data)
        # print("****************** Just Checking **********************")
        # print(list)
        # print(len(data['ChanLow'])) GETS THE LENGTH OF THE DICT
        # print(len(list)) GETS THE LENGTH OF THE LIST
        # print(len(list[0]))
        for a in range(len(list)):
            for b in range(len(list[a])):
                    for x in range(len(data['ChanLow'])): # for in case this length changes with inputs
                        # print(f"this is chanLow = {data['ChanLow'][x]}, and chanHi = {data['ChanHigh'][x]}")
                        chanlo = data['ChanLow'][x]
                        chanHi = data['ChanHigh'][x]
                        self.counter += 1
                        # print(f"This is the list of : {list[b]} YEA BABY counter: {self.counter} chanlo = {chanlo} and CHANHI = {chanHi} COOL")
                        
            