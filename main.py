from getChans import Channels
from getWatts import Watts

watts = Watts()
channels = Channels()
allChans = channels.allChannels
pathTo = channels.path
channels.getChans(allChans)

watts.whatsWatts(path=pathTo, list=allChans)

# have to pass the path and then the allChans to getWatts