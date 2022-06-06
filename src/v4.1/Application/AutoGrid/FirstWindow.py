import sys
# to get the working monitor size
from win32api import GetMonitorInfo, MonitorFromPoint
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtGui import QCursor, QMouseEvent, QFont, QKeySequence, QSyntaxHighlighter, QTextCharFormat, QBrush, QTextCursor
from PyQt5.QtCore import QPoint, pyqtSignal, QRegExp
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect, QEasingCurve
from PyQt5.QtCore import QObject, QMimeData
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QCompleter, QFileDialog, QGraphicsDropShadowEffect
from PyQt5.QtWidgets import QHBoxLayout, QTextEdit, QPlainTextEdit, QShortcut, QScrollArea
from PyQt5.QtWidgets import QLabel, QStackedWidget, QMessageBox
from PyQt5.QtWidgets import QPushButton, QDesktopWidget
from PyQt5.QtWidgets import QVBoxLayout, QScrollBar
from PyQt5.QtWidgets import QWidget, QFrame
from PyQt5.QtCore import Qt, QRect, QSize, QRectF
from PyQt5.QtWidgets import QWidget, QPlainTextEdit, QTextEdit
from PyQt5.QtGui import QColor, QPainter, QTextFormat, QLinearGradient
import os
import ctypes

import TitleBar, config, AutoGrid, Welcome, Month, Zoom, Grid, Person, Extra, Observation, Exam, Loading, Name, Help

class MainWindow(QFrame):
    def __init__(self):
        super(MainWindow, self).__init__()
        # store the main window widget
        global mainWin
        config.mainWin = self
        # set the opacity
        self.setWindowOpacity(1.0)
        # get the current working resolution to account for things like the taskbar
        monitor_info = GetMonitorInfo(MonitorFromPoint((0,0)))
        working_resolution = monitor_info.get("Work")
        workingWidth = working_resolution[2]
        workingHeight = working_resolution[3]
        self.setGeometry(workingWidth/7, 0, workingWidth - (2 * workingWidth / 7), workingHeight)
        # vertical layout
        self.layout = QVBoxLayout()
        self.layout.setSpacing(0)
        # add the title bar
        self.titlebarWidget = TitleBar.MyBar(self)
        self.layout.addWidget(self.titlebarWidget, 5)

        # add drop shadow
        self.shadow = QGraphicsDropShadowEffect()
        self.shadow.setBlurRadius(6)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(2)
        self.shadow.setColor(QColor(0, 0, 0, 200))
        # add a drop shadow before the next thing
        self.dropShadow = QLabel("")
        self.dropShadow.setStyleSheet("""
        background-color: #2E3440;
        border: none;
        
                                        """)
        self.dropShadow.setFixedHeight(1)
        self.dropShadow.setGraphicsEffect(self.shadow)
        self.layout.addWidget(self.dropShadow)

        # create stacked widget layout to put under the title bar
        self.stack = QStackedWidget()
        # store the stack variable
        global stack
        config.stack = self.stack
        # remove the border from the stacked widget
        self.stack.setStyleSheet("border: none;")

        # add a QLabel to the vertical layout
        self.welcome = Welcome.WelcomeScreen(self)
        #self.layout.addWidget(self.welcome, 95)
        
        # add the Qlabel to the stack as index 0
        self.stack.addWidget(self.welcome)

        #------------------------------------------#
        # index 1 needs to be the month selection screen
        self.month = Month.MonthScreen(self)
        self.stack.addWidget(self.month)
        # index 2 needs to be the grid screen
        self.grid = Grid.GridScreen(self)
        self.stack.addWidget(self.grid)
        # index 3 needs to be the zoom screen
        self.zoom = Zoom.ZoomScreen(self)
        self.stack.addWidget(self.zoom)
        # index 4 is in-person attendance
        self.person = Person.PersonScreen(self)
        self.stack.addWidget(self.person)
        # index 5 asks if you've been observed or had any exams
        self.extra = Extra.ExtraScreen(self)
        self.stack.addWidget(self.extra)
        # index 6 is the observation screen
        self.observation = Observation.ObservationScreen(self)
        self.stack.addWidget(self.observation)
        # index 7 is the exams screen
        self.exam = Exam.ExamScreen(self)
        self.stack.addWidget(self.exam)
        # index 8 is the loading screen
        self.loading = Loading.LoadingScreen(self)
        self.stack.addWidget(self.loading)
        # index 9 is the name screen
        self.name = Name.NameScreen(self)
        self.stack.addWidget(self.name)
        # index 10 is the actual naming
        self.help = Help.HelpScreen(self)
        self.stack.addWidget(self.help)
        # start on the latest widget for testing purposes
        #config.stack.setCurrentIndex(7)
        #------------------------------------------#
        # add the stacked widget to the vertical layout
        self.layout.addWidget(self.stack, 95)

        # set the layout
        self.setLayout(self.layout)
        
        # the min height will be 600 x 600
        self.setMinimumSize(config.minSize, config.minSize)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.pressing = False
        self.movingPosition = False
        self.resizingWindow = False
        self.start = QPoint(0, 0)
        self.setStyleSheet("""
            background-color: #2E3440;
            border-style: solid;
            border-width: 1px;
            border-color: #8FBCBB;
                          """)
        
        self.layout.setContentsMargins(config.MARGIN,config.MARGIN,config.MARGIN,config.MARGIN)
        # flags for starting location of resizing window
        self.left = False
        self.right = False
        self.bottom = False
        self.top = False
        self.bl = False
        self.br = False
        self.tl = False
        self.tr = False
        self.top = False
        self.setMouseTracking(True)
        config.app.focusChanged.connect(self.on_focusChanged)
    
    def snapWin(self, direction):
        global rightDown
        global leftDown
        global upDown
        global downDown
        global isMaximized
        
        # start with this so that we can maximize and restore over and over with the up button
        self.showNormal()
        config.isMaximized = False
        # get the current working resolution to account for things like the taskbar
        monitor_info = GetMonitorInfo(MonitorFromPoint((0,0)))
        working_resolution = monitor_info.get("Work")
        workingWidth = working_resolution[2]
        workingHeight = working_resolution[3]
        # determine if the taskbar is present by comparing the normal height to the working height
        isTaskbar = True
        difference = 100000
        for i in range(0, QDesktopWidget().screenCount()):
            if workingHeight == QDesktopWidget().screenGeometry(i).height():
                isTaskbar = False
                break
            # store the smallest difference to determine the correct difference due to the taskbar
            elif abs(QDesktopWidget().screenGeometry(i).height() - workingHeight) < difference:
                difference = QDesktopWidget().screenGeometry(i).height() - workingHeight

        # if the taskbar is present then use the working height
        if isTaskbar == True:
            workingWidth = QDesktopWidget().screenGeometry(self).width()
            workingHeight = QDesktopWidget().screenGeometry(self).height() - difference
        # if the taskbar is not present then just use the normal width and height
        else:
            workingWidth = QDesktopWidget().screenGeometry(self).width()
            workingHeight = QDesktopWidget().screenGeometry(self).height()
        
        monitor = QDesktopWidget().screenGeometry(self)
        self.move(monitor.left(), monitor.top())

        # middle window from right
        if direction == "left" and config.rightDown == True:
            self.setGeometry(monitor.left() + workingWidth/4, monitor.top(), workingWidth/2, workingHeight)
            # set the m all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # middle window from left
        elif direction == "right" and config.leftDown == True:
            self.setGeometry(monitor.left() + workingWidth/4, monitor.top(), workingWidth/2, workingHeight)
            # set the m all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap the window right
        elif direction == "right" and config.downDown == False and config.upDown == False:
            self.setGeometry(monitor.left() + workingWidth/2, monitor.top(), workingWidth/2, workingHeight)
            # set the right to true and the others to false
            config.rightDown = True
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap bottom right from bottom
        elif direction == "right" and config.downDown == True and config.upDown == False:
            self.setGeometry(monitor.left() + workingWidth/2, monitor.top() + workingHeight/2, workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap bottom right from right
        elif direction == "bottom" and config.leftDown == False and config.rightDown == True:
            self.setGeometry(monitor.left() + workingWidth/2, monitor.top() + workingHeight/2, workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False

        # snap bottom left from bottom
        elif direction == "left" and config.downDown == True and config.upDown == False:
            self.setGeometry(monitor.left(), monitor.top() + workingHeight/2, workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap bottom left from left
        elif direction == "bottom" and config.leftDown == True and config.rightDown == False:
            self.setGeometry(monitor.left(), monitor.top() + workingHeight/2, workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap top left from top
        elif direction == "left" and config.downDown == False and config.upDown == True:
            self.setGeometry(monitor.left(), monitor.top(), workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # maximize
        elif direction == "top" and config.upDown == True:
            # click the max button
            self.setGeometry(monitor.left(), monitor.top(), workingWidth, workingHeight)
            config.isMaximized = True
            #self.layout.itemAt(0).widget().btn_max_clicked()
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap top left from left
        elif direction == "top" and config.leftDown == True and config.rightDown == False:
            self.setGeometry(monitor.left(), monitor.top(), workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap top right from top
        elif direction == "right" and config.downDown == False and config.upDown == True:
            self.setGeometry(monitor.left() + workingWidth / 2, monitor.top(), workingWidth/2, workingHeight/2)
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False
        
        # snap top right from right
        elif direction == "top" and config.leftDown == False and config.rightDown == True:
            self.setGeometry(monitor.left() + workingWidth / 2, monitor.top(), workingWidth/2, workingHeight/2)   
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False

        # snap left
        elif direction == "left" and config.downDown == False and config.upDown == False:
            self.setGeometry(monitor.left(), monitor.top(), workingWidth/2, workingHeight)
            # set left to true and others to false
            config.leftDown = True
            config.rightDown = False
            config.downDown = False
            config.upDown = False

        # snap up
        elif direction == "top" and config.leftDown == False and config.rightDown == False:
            self.setGeometry(monitor.left(), monitor.top(), workingWidth, workingHeight / 2)
            # set up to True and all others to false
            config.upDown = True
            config.leftDown = False
            config.rightDown = False
            config.downDown = False
        
        # minimize
        elif direction == "bottom" and config.downDown == True:
            # click the min button
            self.layout.itemAt(0).widget().btn_min_clicked()
            # set all to false
            config.rightDown = False
            config.leftDown = False
            config.downDown = False
            config.upDown = False

        # snap down
        elif direction == "bottom" and config.leftDown == False and config.rightDown == False:
            self.setGeometry(monitor.left(), monitor.top() + workingHeight / 2, workingWidth, workingHeight / 2)
            # set Down to True and all others to false
            config.downDown = True
            config.upDown = False
            config.leftDown = False
            config.rightDown = False     
        
        mainPosition = config.mainWin.mapToGlobal(QPoint(0,config.mainWin.height()))
        self.updateSize()

    def on_focusChanged(self, old, new):
        # set the opacity to 1 if not focused
        if self.isActiveWindow():
            self.setWindowOpacity(0.98)
        else:
            self.setWindowOpacity(1.0)
    
    def mousePressEvent(self, event):
        pos = event.pos()
        # set pressing to true
        self.pressing = True
        if config.isMaximized == False:
            # if they clicked on the edge then we need to change pressing to true and resizingWindow to
            # true and we need to change the cursor shape
            # top left
            if pos.x() <= 8 and pos.y() <= 8:
                self.resizingWindow = True
                self.start = event.pos()
                self.tl = True
            # top right
            elif pos.x() >= self.width() - 8 and pos.y() <= 8:
                self.resizingWindow = True
                self.start = event.pos()
                self.tr = True
            # top
            elif pos.y() <= 8 and pos.x() > 8 and pos.x() < self.width() - 8:
                self.resizingWindow = True
                self.start = event.pos().y()
                self.top = True     
            elif pos.y() >= self.height() - 8 and pos.x() <= 8 and pos.y() > 8:
                self.resizingWindow = True
                self.start = event.pos()
                self.bl = True
            elif pos.x() <= 8 and pos.y() > 8:
                self.resizingWindow = True
                self.start = event.pos().x()
                self.left = True   
            elif pos.x() >= self.width() - 8 and pos.y() >= self.height() - 8:
                self.resizingWindow = True
                self.start = event.pos()
                self.br = True    
            elif pos.x() >= self.width() - 8 and pos.y() > 8:
                self.resizingWindow = True
                self.start = event.pos().x()
                self.right = True              
            elif pos.x() > 8 and pos.x() < self.width() - 8 and pos.y() >= self.height() - 8:
                self.resizingWindow = True
                self.start = event.pos().y()
                self.bottom = True   
  
    def mouseMoveEvent(self, event):
        pos = event.pos()
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        if config.isMaximized == False:
            # top left
            if pos.x() <= 10 and pos.y() <= 10:
                QApplication.setOverrideCursor(Qt.SizeFDiagCursor)
            # top right
            elif pos.x() >= self.width() - 8 and pos.y() <= 8:
                QApplication.setOverrideCursor(Qt.SizeBDiagCursor)
            # top
            elif pos.y() <= 5 and pos.x() > 5 and pos.x() < self.width() - 5:
                QApplication.setOverrideCursor(Qt.SizeVerCursor)
            # bottom left
            elif pos.y() >= self.height() - 8 and pos.x() <= 8:
                QApplication.setOverrideCursor(Qt.SizeBDiagCursor)
            # bottom right
            elif pos.x() >= self.width() - 8 and pos.y() >= self.height() - 8:
                QApplication.setOverrideCursor(Qt.SizeFDiagCursor)
            # bottom
            elif pos.x() > 0 and pos.x() < self.width() - 8 and pos.y() >= self.height() - 8:
                QApplication.setOverrideCursor(Qt.SizeVerCursor)
            # left
            elif pos.x() <= 5 and pos.y() > 5:
                QApplication.setOverrideCursor(Qt.SizeHorCursor)
            # right
            elif pos.x() >= self.width() - 5 and pos.y() > 5:
                QApplication.setOverrideCursor(Qt.SizeHorCursor)
            else:
                QApplication.setOverrideCursor(Qt.ArrowCursor)
            self.updateSize()


        # if they are resizing
        # need to subtract the movement from the width/height 
        # but I also need to account for if they are resizing horizontally from the left or
        # vertically from the top because I need to shift the window to the right/down the same amount
        if self.pressing and self.resizingWindow:
            # resize from the top
            if self.top == True:
                # resize from the top
                if self.height() - event.pos().y() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y() + event.pos().y(), self.width(), self.height() - event.pos().y())
            # resize from the top left
            if self.tl == True:
                # move both dimensions if both boundaries are okay
                if self.width() - event.pos().x() >= config.minSize and self.height() - event.pos().y() >= config.minSize:
                    self.setGeometry(self.pos().x() + event.pos().x(), self.pos().y() + event.pos().y(), self.width() - event.pos().x(), self.height() - event.pos().y())
                # move only top if width is already at its smallest
                elif self.height() - event.pos().y() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y() + event.pos().y(), self.width(), self.height() - event.pos().y())
                # move only left if height is at its smallest
                elif self.width() - event.pos().x() > config.minSize:
                    self.setGeometry(self.pos().x() + event.pos().x(), self.pos().y(), self.width() - event.pos().x(), self.height())
            
            # resize top right
            if self.tr == True:
                pos = event.pos().x() 
                # top right
                if self.height() - event.pos().y() >= config.minSize and self.width() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y() + event.pos().y(), pos, self.height() - event.pos().y())

                # resize from the top
                elif self.height() - event.pos().y() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y() + event.pos().y(), self.width(), self.height() - event.pos().y())
                elif self.width() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y(), pos, self.height()) 

            # resize from the left to the right
            if self.left == True:
                # resize from the left
                if self.width() - event.pos().x() > config.minSize:
                    self.setGeometry(self.pos().x() + event.pos().x(), self.pos().y(), self.width() - event.pos().x(), self.height())
            # resize from the right
            if self.right == True:
                pos = event.pos().x()
                if self.width() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y(), pos, self.height()) 
            # resize from the bottom
            if self.bottom == True:
                pos = event.pos().y()
                if self.height() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y(), self.width(), pos) 
            # resize from the bottom right
            if self.br == True:
                pos = event.pos()
                if self.height() >= config.minSize and self.width() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y(), pos.x(), pos.y()) 
            # resize from the bottom left
            if self.bl == True:
                pos = event.pos().y()
                if self.width() - event.pos().x() > config.minSize and self.height() >= config.minSize:
                    self.setGeometry(self.pos().x() + event.pos().x(), self.pos().y(), self.width() - event.pos().x(), pos)
                elif self.height() >= config.minSize:
                    self.setGeometry(self.pos().x(), self.pos().y(), self.width(), pos) 
                elif self.width() - event.pos().x() > config.minSize:
                    self.setGeometry(self.pos().x() + event.pos().x(), self.pos().y(), self.width() - event.pos().x(), self.height())
            
    # if the mouse button is released then tag pressing as false
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            return
        self.pressing = False
        self.movingPosition = False
        self.resizingWindow = False
        self.left = False
        self.right = False
        self.bottom = False
        self.bl = False
        self.br = False
        self.tr = False
        self.tl = False
        self.top = False

    def updateSize(self):
        # change the text of the welcome screen
        font = QFont()
        font.setFamily("Serif")
        font.setFixedPitch( True )
        font.setPointSize( self.width() / 20 )
        config.welcomeScreen.welcomeText.setFont(font)
        font2 = QFont()
        font2.setFamily("Serif")
        font2.setFixedPitch( True )
        font2.setPointSize( self.width() / 45)
        config.welcomeScreen.pressStart.setFont(font2)
        # change the size of the welcome start button and its font
        config.welcomeScreen.welcomeButton.setFixedHeight(self.width() / 12)
        config.welcomeScreen.welcomeButton.setFixedWidth(self.width() / 3)
        # set the font of teh button
        font3 = QFont()
        font3.setFamily("Serif")
        font3.setFixedPitch( True )
        font3.setPointSize( self.width() / 35 )
        config.welcomeScreen.welcomeButton.setFont(font3)
        # set the font size of the month screen
        monthFont = QFont()
        monthFont.setFamily("Serif")
        monthFont.setFixedPitch( True )
        monthFont.setPointSize( self.width() / 30 )
        config.monthScreen.monthText.setFont(monthFont)
        # set the size of the month dropdown menu in the month screen
        # set the size of the month select dropdown menu
        config.monthScreen.monthSelect.setFixedHeight(self.width() / 22)
        config.monthScreen.monthSelect.setFixedWidth(self.width() / 3)
        # set the font sizeof the month dropdown  
        monthFont = QFont()
        monthFont.setFamily("Serif")
        monthFont.setFixedPitch( True )
        monthFont.setPointSize( self.width() / 40 )
        config.monthScreen.monthSelect.setFont(monthFont)    
        # set the font size and button size of the month continue button
        # set the size of the button
        config.monthScreen.monthButton.setFixedHeight(self.width() / 12)
        config.monthScreen.monthButton.setFixedWidth(self.width() / 3)
        
        # set the font of teh button
        monthButtonFont = QFont()
        monthButtonFont.setFamily("Serif")
        monthButtonFont.setFixedPitch( True )
        monthButtonFont.setPointSize( self.width() / 35 )
        config.monthScreen.monthButton.setFont(monthButtonFont)

        # update the zoom text font size
        zoomFont = QFont()
        zoomFont.setFamily("Serif")
        zoomFont.setFixedPitch( True )
        zoomFont.setPointSize( self.width() / 32 )
        config.zoomScreen.zoomText.setFont(zoomFont)
        config.zoomScreen.numFiles.setFixedHeight(self.width() / 24)
        config.zoomScreen.numFiles.setFixedWidth(self.width() / 3)

        # update numFiles QLabel font and size
        numFilesFont = QFont()
        numFilesFont.setFamily("Serif")
        numFilesFont.setFixedPitch( True )
        numFilesFont.setPointSize( self.width() / 53 )
        config.zoomScreen.numFiles.setFont(numFilesFont)    

        # update browse button size and font
        browseButtonFont = QFont()
        browseButtonFont.setFamily("Serif")
        browseButtonFont.setFixedPitch( True )
        browseButtonFont.setPointSize( self.width() / 40 )
        config.zoomScreen.fileSelect.setFont(browseButtonFont)    
        config.zoomScreen.fileSelect.setFixedHeight(self.width() / 22)
        config.zoomScreen.fileSelect.setFixedWidth(self.width() / 6)

        # update zoom continue button font and size
        zoomButtonFont = QFont()
        zoomButtonFont.setFamily("Serif")
        zoomButtonFont.setFixedPitch( True )
        zoomButtonFont.setPointSize( self.width() / 35 )
        config.zoomScreen.zoomButton.setFont(zoomButtonFont)
        config.zoomScreen.zoomButton.setFixedHeight(self.width() / 12)
        config.zoomScreen.zoomButton.setFixedWidth(self.width() / 3)

        # update grid text font
        gridFont = QFont()
        gridFont.setFamily("Serif")
        gridFont.setFixedPitch( True )
        gridFont.setPointSize( self.width() / 32 )
        config.gridScreen.gridText.setFont(gridFont)

        # update hte size of the browse button and font
        config.gridScreen.fileSelect.setFixedHeight(self.width() / 22)
        config.gridScreen.fileSelect.setFixedWidth(self.width() / 6)
        browseButtonFont = QFont()
        browseButtonFont.setFamily("Serif")
        browseButtonFont.setFixedPitch( True )
        browseButtonFont.setPointSize( self.width() / 40 )
        config.gridScreen.fileSelect.setFont(browseButtonFont)    

        # font that displays the file name and size
        numFilesFont = QFont()
        numFilesFont.setFamily("Serif")
        numFilesFont.setFixedPitch( True )
        numFilesFont.setPointSize( self.width() / 90 )
        config.gridScreen.numFiles.setFont(numFilesFont)    
        config.gridScreen.numFiles.setFixedHeight(self.width() / 24)
        config.gridScreen.numFiles.setFixedWidth(self.width() / 3)

        # grid button font and size
        gridButtonFont = QFont()
        gridButtonFont.setFamily("Serif")
        gridButtonFont.setFixedPitch( True )
        gridButtonFont.setPointSize( self.width() / 35 )
        config.gridScreen.gridButton.setFont(gridButtonFont)
        config.gridScreen.gridButton.setFixedHeight(self.width() / 12)
        config.gridScreen.gridButton.setFixedWidth(self.width() / 3)

        # update the text of the in-person screen
        gridFont = QFont()
        gridFont.setFamily("Serif")
        gridFont.setFixedPitch( True )
        gridFont.setPointSize( self.width() / 32 )
        config.personScreen.personText.setFont(gridFont)

        # browse button in-person
        config.personScreen.fileSelect.setFixedHeight(self.width() / 22)
        config.personScreen.fileSelect.setFixedWidth(self.width() / 6)
        browseButtonFont = QFont()
        browseButtonFont.setFamily("Serif")
        browseButtonFont.setFixedPitch( True )
        browseButtonFont.setPointSize( self.width() / 40 )
        config.personScreen.fileSelect.setFont(browseButtonFont) 

        # qlabel in-person
        config.personScreen.numFiles.setFixedHeight(self.width() / 24)
        config.personScreen.numFiles.setFixedWidth(self.width() / 3)
        numFilesFont = QFont()
        numFilesFont.setFamily("Serif")
        numFilesFont.setFixedPitch( True )
        numFilesFont.setPointSize( self.width() / 90 )
        config.personScreen.numFiles.setFont(numFilesFont)    

        # in-person continue button
        config.personScreen.personButton.setFixedHeight(self.width() / 12)
        config.personScreen.personButton.setFixedWidth(self.width() / 3)
        personButtonFont = QFont()
        personButtonFont.setFamily("Serif")
        personButtonFont.setFixedPitch( True )
        personButtonFont.setPointSize( self.width() / 35 )
        config.personScreen.personButton.setFont(personButtonFont)

        # observed font
        observedFont = QFont()
        observedFont.setFamily("Serif")
        observedFont.setFixedPitch( True )
        observedFont.setPointSize( self.width() / 40 )
        config.extraScreen.observedText.setFont(observedFont)

        # observed dropdown
        config.extraScreen.observedSelect.setFixedHeight(self.width() / 22)
        config.extraScreen.observedSelect.setFixedWidth(self.width() / 15)
        observedDropdownFont = QFont()
        observedDropdownFont.setFamily("Serif")
        observedDropdownFont.setFixedPitch( True )
        observedDropdownFont.setPointSize( self.width() / 40 )
        config.extraScreen.observedSelect.setFont(observedDropdownFont)    

        # extra button font and size
        config.extraScreen.extraButton.setFixedHeight(self.width() / 12)
        config.extraScreen.extraButton.setFixedWidth(self.width() / 3)
        
        # set the font of teh button
        extraButtonFont = QFont()
        extraButtonFont.setFamily("Serif")
        extraButtonFont.setFixedPitch( True )
        extraButtonFont.setPointSize( self.width() / 35 )
        config.extraScreen.extraButton.setFont(extraButtonFont)

        # exam text
        examFont = QFont()
        examFont.setFamily("Serif")
        examFont.setFixedPitch( True )
        examFont.setPointSize( self.width() / 45 )
        config.extraScreen.examText.setFont(examFont)

        # exam dropdown
        config.extraScreen.examSelect.setFixedHeight(self.width() / 22)
        config.extraScreen.examSelect.setFixedWidth(self.width() / 15)
        examDropdownFont = QFont()
        examDropdownFont.setFamily("Serif")
        examDropdownFont.setFixedPitch( True )
        examDropdownFont.setPointSize( self.width() / 40 )
        config.extraScreen.examSelect.setFont(examDropdownFont) 

        # observation title
        observedFont = QFont()
        observedFont.setFamily("Serif")
        observedFont.setFixedPitch( True )
        observedFont.setPointSize( self.width() / 44 )
        config.observationScreen.observedText.setFont(observedFont) 

        # first horizontal layout stuff
        # observer 1 | name1 | last1
        # observer size
        config.observationScreen.observer1.setFixedHeight(self.width() / 24)
        config.observationScreen.observer1.setFixedWidth(self.width() / 3)
        config.observationScreen.observer2.setFixedHeight(self.width() / 24)
        config.observationScreen.observer2.setFixedWidth(self.width() / 3)
        config.observationScreen.observer3.setFixedHeight(self.width() / 24)
        config.observationScreen.observer3.setFixedWidth(self.width() / 3)
        config.observationScreen.observer4.setFixedHeight(self.width() / 24)
        config.observationScreen.observer4.setFixedWidth(self.width() / 3)
        # observer font
        observerNumberFont = QFont()
        observerNumberFont.setFamily("Serif")
        observerNumberFont.setFixedPitch( True )
        observerNumberFont.setPointSize( self.width() / 53 )
        config.observationScreen.observer1.setFont(observerNumberFont)   
        config.observationScreen.observer2.setFont(observerNumberFont)   
        config.observationScreen.observer3.setFont(observerNumberFont)   
        config.observationScreen.observer4.setFont(observerNumberFont)   

        # first name size
        config.observationScreen.name1.setFixedHeight(self.width() / 24)
        config.observationScreen.name1.setFixedWidth(self.width() / 5)
        config.observationScreen.name2.setFixedHeight(self.width() / 24)
        config.observationScreen.name2.setFixedWidth(self.width() / 5)
        config.observationScreen.name3.setFixedHeight(self.width() / 24)
        config.observationScreen.name3.setFixedWidth(self.width() / 5)
        config.observationScreen.name4.setFixedHeight(self.width() / 24)
        config.observationScreen.name4.setFixedWidth(self.width() / 5)
        # last name size
        config.observationScreen.last1.setFixedHeight(self.width() / 24)
        config.observationScreen.last1.setFixedWidth(self.width() / 5)
        config.observationScreen.last2.setFixedHeight(self.width() / 24)
        config.observationScreen.last2.setFixedWidth(self.width() / 5)
        config.observationScreen.last3.setFixedHeight(self.width() / 24)
        config.observationScreen.last3.setFixedWidth(self.width() / 5)
        config.observationScreen.last4.setFixedHeight(self.width() / 24)
        config.observationScreen.last4.setFixedWidth(self.width() / 5)

        # first and last font
        textFont = QFont()
        textFont.setFamily("Serif")
        textFont.setFixedPitch( True )
        textFont.setPointSize( self.width() / 53 )
        
        config.observationScreen.name1.setFont(textFont)    
        config.observationScreen.last1.setFont(textFont)    
        config.observationScreen.name2.setFont(textFont)    
        config.observationScreen.last2.setFont(textFont)    
        config.observationScreen.name3.setFont(textFont)    
        config.observationScreen.last3.setFont(textFont)    
        config.observationScreen.name4.setFont(textFont)    
        config.observationScreen.last4.setFont(textFont)

        # continue button size
        config.observationScreen.observationButton.setFixedHeight(self.width() / 12)
        config.observationScreen.observationButton.setFixedWidth(self.width() / 3)
        
        # set the font of teh button
        observationButtonFont = QFont()
        observationButtonFont.setFamily("Serif")
        observationButtonFont.setFixedPitch( True )
        observationButtonFont.setPointSize( self.width() / 35 )
        config.observationScreen.observationButton.setFont(observationButtonFont)

        # Exam title font
        examFont = QFont()
        examFont.setFamily("Serif")
        examFont.setFixedPitch( True )
        examFont.setPointSize( self.height() / 20 )
        config.examScreen.examText.setFont(examFont)

        # Button sizes & font
        monthButtonFont = QFont()
        monthButtonFont.setFamily("Serif")
        monthButtonFont.setFixedPitch( True )
        monthButtonFont.setPointSize( self.height() / 40 )
        for i in range(0, 31):
            config.examScreen.buttonArr[i].setFixedHeight(self.height() / 18)
            config.examScreen.buttonArr[i].setFixedWidth(self.width() / 18)
            config.examScreen.buttonArr[i].setFont(monthButtonFont)   
            config.nameScreen.buttonArr[i].setFixedHeight(self.height() / 18)
            config.nameScreen.buttonArr[i].setFixedWidth(self.width() / 18)
            config.nameScreen.buttonArr[i].setFont(monthButtonFont)   
        
        # exam button
        examButtonFont = QFont()
        examButtonFont.setFamily("Serif")
        examButtonFont.setFixedPitch( True )
        examButtonFont.setPointSize( self.width() / 35 )
        config.examScreen.examButton.setFont(examButtonFont)
        config.examScreen.examButton.setFixedHeight(self.width() / 12)
        config.examScreen.examButton.setFixedWidth(self.width() / 3)

        # loading screen font
        loadingFont = QFont()
        loadingFont.setFamily("Serif")
        loadingFont.setFixedPitch( True )
        loadingFont.setPointSize( self.width() / 27 )
        config.loadingScreen.loadingText.setFont(loadingFont)

        # name help text font
        nameTextFont = QFont()
        nameTextFont.setFamily("Serif")
        nameTextFont.setFixedPitch( True )
        nameTextFont.setPointSize( self.width() / 50 )
        config.nameScreen.nameText.setFont(nameTextFont)

        # name help infotext
        infoTextFont = QFont()
        infoTextFont.setFamily("Serif")
        infoTextFont.setFixedPitch( True )
        infoTextFont.setPointSize( self.width() / 110 )
        config.nameScreen.infoText.setFont(infoTextFont)

        # Button sizes & font
        monthButtonFont2 = QFont()
        monthButtonFont2.setFamily("Serif")
        monthButtonFont2.setFixedPitch( True )
        monthButtonFont2.setPointSize( self.height() / 55 )
        for i in range(0, 31):
            config.nameScreen.buttonArr[i].setFixedHeight(self.height() / 30)
            config.nameScreen.buttonArr[i].setFixedWidth(self.width() / 30)
            config.nameScreen.buttonArr[i].setFont(monthButtonFont2)   
        
        # name dropdown
        config.nameScreen.scenario.setFixedHeight(self.height() / 10)
        config.nameScreen.scenario.setFixedWidth(self.width() / 3)
        scenarioFont = QFont()
        scenarioFont.setFamily("Serif")
        scenarioFont.setFixedPitch( True )
        scenarioFont.setPointSize( self.width() / 80 )
        config.nameScreen.scenario.setFont(scenarioFont) 

        # checkbox font
        checkFont = QFont()
        checkFont.setFamily("Serif")
        checkFont.setFixedPitch( True )
        checkFont.setPointSize( self.width() / 90 )
        config.nameScreen.reviewCheck.setFont(checkFont)

        # name continue button
        config.nameScreen.nameButton.setFixedHeight(self.width() / 15)
        config.nameScreen.nameButton.setFixedWidth(self.width() / 6)
        config.nameScreen.backButton.setFixedHeight(self.width() / 15)
        config.nameScreen.backButton.setFixedWidth(self.width() / 6)

        nameButtonFont = QFont()
        nameButtonFont.setFamily("Serif")
        nameButtonFont.setFixedPitch( True )
        nameButtonFont.setPointSize( self.width() / 44)
        config.nameScreen.nameButton.setFont(nameButtonFont)
        config.nameScreen.backButton.setFont(nameButtonFont)

        # help screen text
        helpTextFont = QFont()
        helpTextFont.setFamily("Serif")
        helpTextFont.setFixedPitch( True )
        helpTextFont.setPointSize( self.width() / 50 )
        config.helpScreen.helpText.setFont(helpTextFont)

        # help screen button
        config.helpScreen.helpButton.setFixedHeight(self.width() / 12)
        config.helpScreen.helpButton.setFixedWidth(self.width() / 3)
        helpButtonFont = QFont()
        helpButtonFont.setFamily("Serif")
        helpButtonFont.setFixedPitch( True )
        helpButtonFont.setPointSize( self.width() / 35 )
        config.helpScreen.helpButton.setFont(helpButtonFont)

        # help font on welcome screen
        helpFont = QFont()
        helpFont.setFamily("Serif")
        helpFont.setFixedPitch( True )
        helpFont.setPointSize( self.width() / 100 )
        config.welcomeScreen.helpText.setFont(helpFont)

        # guide button on welcome screen
        config.welcomeScreen.helpButton.setFixedHeight(self.width() / 40)
        config.welcomeScreen.helpButton.setFixedWidth(self.width() / 12)
        helpButtonFont2 = QFont()
        helpButtonFont2.setFamily("Serif")
        helpButtonFont2.setFixedPitch( True )
        helpButtonFont2.setPointSize( self.width() / 80 )
        config.welcomeScreen.helpButton.setFont(helpButtonFont2)