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
daWatts = watts.temp
getTotalWatts = totalWatts.totalWattsPerLine

watts.whatsWatts(path=pathTo, list=allChans, daWatts=daWatts)
print(daWatts)

totalWatts.getDaTotals(daWatts, totalWattsPerLine=getTotalWatts)
