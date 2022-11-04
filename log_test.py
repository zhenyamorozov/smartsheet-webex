import logging 
import io

logger = logging.getLogger(__name__)
# set log level to DEBUG
# Level DEBUG - to be output in a full log (text file attached to Webex message)
# Level INFO - to be output to a brief log (the Webex message itself)
logger.setLevel(logging.DEBUG)

briefLogString = io.StringIO() 
briefLogHandler = logging.StreamHandler(briefLogString) 
briefLogHandler.setLevel(logging.INFO)
logger.addHandler(briefLogHandler)

fullLogString = io.StringIO() 
fullLogHandler = logging.StreamHandler(fullLogString) 
fullLogHandler.setLevel(logging.DEBUG)
logger.addHandler(fullLogHandler)

consoleLogHandler = logging.StreamHandler() 
consoleLogHandler.setLevel(logging.DEBUG)
logger.addHandler(consoleLogHandler)



logger.error("there was an error")
logger.info("Message to the brief log")
logger.debug("message to the full log")

print("BRIEF:")
print(briefLogString.getvalue())

print("FULL:")
print(fullLogString.getvalue())


briefLogString.close()
fullLogString.close()
