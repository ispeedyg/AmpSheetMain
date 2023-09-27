import pandas as pd 





class Watts:
    def __init__(self) -> None:
        self.chanLow = int
        self.chanHi = int
        
        
        
    def whatsWatts(self, path, list):
        data = pd.read_excel(path, 'Watts').to_dict()
        print(data)
        print("****************************************")
        print(list)
        