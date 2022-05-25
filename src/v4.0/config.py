appName = "AutoGrid"
# make the resolution global variables
screen_resolution = 0
width = 0
height = 0
key = ''
res = {}
res["1920x1080"] = [1920/2, 0] # full hd
res["2560x1440"] = [2560/2, 0] # wqhd
res["3440x1440"] = [3440/2, 0] # ultrawide
res["3840x2160"] = [3840/2, 0] # 4k
focused = False # variable to track if the gui is focused so it knows to track typing or not

# variable to track the margins used on the main layout
MARGIN = 5

# variable to allow going back to previous size after maximizing
isMaximized = False

# variables to store the mainwindow and title bar
app = None
mainWin = None
titleBar = None
welcomeScreen = None
# store the widget stack
stack = None
# store the month screen class
monthScreen = None
# store the zoom screen
zoomScreen = None
# store the grid screen
gridScreen = None
# store in person attendance screen
personScreen = None
# store the extra questions screen
extraScreen = None
# observation screen
observationScreen = None
# exam screen
examScreen = None
# loading screen
loadingScreen = None
# name help screen
nameScreen = None
# help screen
helpScreen = None

# variable to be able to snap to sides and corners
leftDown = False
upDown = False
downDown = False
rightDown = False

# wait time
waitTime = 300
flashNumber = 3

# variables for color settings
bracketColor = "#D08770"
keywordColor = "#81A1C1"
parenColor = "#EBCB8B"
braceColor = "#D08770"
functionColor = "#88C0D0"
commentColor = "#4C566A"
textColor = "#D8DEE9"
accentColor = "#8FBCBB"
accentColor2 = "#A3BE8C"
numberColor = "#BF616A"
backgroundColor = "#2E3440"
lineNumberColor = "#8FBCBB"
selectionColor = "#4C566A"
curLineColor = "#3B4252"
selectionTextColor = "#D8DEE9"
classColor = "#A3BE8C"
operatorColor = "#EBCB8B"
unclosedString = "#BF616A"

fontSize = 14