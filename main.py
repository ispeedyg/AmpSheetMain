from getChans import Channels
from getSpares import GetSpareCircuits
from getWatts import Watts
from lineWatts import TotalWatts
 

newList = []
watts = Watts()
channels = Channels()
spareCirc = GetSpareCircuits()
totalWatts = TotalWatts()

allChans = channels.allChannels # just a thought for this try making it a [] to see if the return really works
pathTo = channels.path
channels.getChans(allChans)
rowLength = channels.rowLength
# print(f"\nOriginal list of channels: \n{allChans}\n")
# print(f"The highest rowlength is : {rowLength}\n")

spareCirc.spares(hiLenRows=rowLength, origList=allChans)
daWatts = watts.temp
getTotalWatts = totalWatts.totalWattsPerLine

watts.whatsWatts(path=pathTo, list=allChans)


# totalWatts.getDaTotals(daWatts, totalWattsPerLine=getTotalWatts)
