from getChans import Channels
from getWatts import Watts
 

newList = []
watts = Watts()
channels = Channels()
allChans = channels.allChannels
pathTo = channels.path
channels.getChans(allChans)

watts.whatsWatts(path=pathTo, list=allChans)

