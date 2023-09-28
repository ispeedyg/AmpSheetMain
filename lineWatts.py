import pandas as pd 

class TotalWatts:
    def __init__(self) -> None:
        self.lineWatts = 0
        self.totalLineWatts = []
        self.totalWattsPerLine = []
        
        
    def getDaTotals(self, daWatts, totalWattsPerLine):
        df = pd.DataFrame(daWatts)
        print(df)