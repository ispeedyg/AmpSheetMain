from getChans import Channels
from getWatts import Watts
newList = []
watts = Watts()
channels = Channels()
allChans = channels.allChannels
pathTo = channels.path
channels.getChans(allChans)
newList = [subList[0] for subList in allChans]
print(newList)
watts.whatsWatts(path=pathTo, list=allChans)

