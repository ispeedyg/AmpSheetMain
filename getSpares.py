import pandas as pd   

class GetSpareCircuits:
    def __init__(self) -> None:
        self.spare = 0
        self.start = 0
        self.topOfCircuit = 1 # has to be one more to be inclusive in range
        self.MAXROWS = int
        self.circuitCounter = -1
        
    def spares(self, hiLenRows, origList):
        self.MAXROWS = hiLenRows
        # print(f"the maximun rows columns is: {self.MAXROWS}\n") # passes through just fine 
        # print(f"original list from get channels is: {origList}\n") # passes through just fine list in list though
        df = pd.DataFrame(origList)
        
        for i, row in df.iterrows():
            self.circuitCounter += 1
            # print(i, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7])
            # print(df.loc[[self.circuitCounter]]) # Wrap your index value with an addition pair of square brackets prints the entire row
            # if not df.loc[[self.circuitCounter]].empty:
            #     print(f"line {self.circuitCounter} has a value")
            # else:
            #     print(f"line {self.circuitCounter} is a spare line")
            # above is an almost but nan must be considered a value. But it almost got something
            # think i found the answer i need here https://stackoverflow.com/questions/29530232/how-to-check-if-any-value-is-nan-in-a-pandas-dataframe
            
                
                
                
                