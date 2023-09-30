from getChans import Channels
from getWatts import Watts
from lineWatts import TotalWatts
 

newList = []
watts = Watts()
channels = Channels()
totalWatts = TotalWatts()
allChans = channels.allChannels

pathTo = channels.path
channels.getChans(allChans)
rowLength = channels.rowLength
print(f"Original list of channels: {allChans}\n")
print(f"The highest rowlength is : {rowLength}\n")
daWatts = watts.temp
getTotalWatts = totalWatts.totalWattsPerLine

watts.whatsWatts(path=pathTo, list=allChans, daWatts=daWatts)


totalWatts.getDaTotals(daWatts, totalWattsPerLine=getTotalWatts)
