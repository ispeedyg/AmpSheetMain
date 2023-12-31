import pandas as pd 

class Watts:
    def __init__(self) -> None:
        self.chanLow = int
        self.chanHi = int
        self.watts = int
        self.temp = []
        
     
    def whatsWatts(self, path, list):
        data = pd.read_excel(path, 'Watts').to_dict()
        df = pd.DataFrame(list)
        # print(df)
        for a in range(48):           
            for x in range(len(data['ChanLow'])): # for in case this length changes with inputs
                # print(f"this is chanLow = {data['ChanLow'][x]}, and chanHi = {data['ChanHigh'][x]+1}")
                chanlo = data['ChanLow'][x]
                chanHi = data['ChanHigh'][x] + 1 # should make the 9z inclusive in range by adding 1 which will be like the hundreds minus 1
        
                for b in range(len(df.loc[a])):
                    # print(len(df.loc[a]))
                    if df.loc[a][b] in range(chanlo, chanHi): 
                        # print(f"im in there: {df.loc[a][b]}")
                        self.watts = data['Wattage'][x]
                        # print(self.watts)
                        self.temp.append(a)
                        self.temp.append(self.watts)
        print(self.temp)   
        return(self.temp)
            