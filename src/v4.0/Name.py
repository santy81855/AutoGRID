import sys
# to get the working monitor size
from win32api import GetMonitorInfo, MonitorFromPoint
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtGui import QCursor, QMouseEvent, QFont, QKeySequence, QSyntaxHighlighter, QTextCharFormat, QBrush, QTextCursor
from PyQt5.QtCore import QPoint, pyqtSignal, QRegExp
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect, QEasingCurve
from PyQt5.QtCore import QObject, QMimeData
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QCompleter, QFileDialog, QGraphicsDropShadowEffect, QComboBox, QCheckBox
from PyQt5.QtWidgets import QHBoxLayout, QTextEdit, QPlainTextEdit, QShortcut, QScrollArea
from PyQt5.QtWidgets import QLabel, QStackedWidget, QMessageBox
from PyQt5.QtWidgets import QPushButton, QDesktopWidget, QSizePolicy
from PyQt5.QtWidgets import QVBoxLayout, QScrollBar
from PyQt5.QtWidgets import QWidget, QFrame
from PyQt5.QtCore import Qt, QRect, QSize, QRectF
from PyQt5.QtWidgets import QWidget, QPlainTextEdit, QTextEdit
from PyQt5.QtGui import QColor, QPainter, QTextFormat, QLinearGradient
from PyQt5 import QtTest
import os
import ctypes

import TitleBar, config, AutoGrid, Welcome, FirstWindow, Month

class Check(QCheckBox):
    def __init__(self, parent, text):
        super(Check, self).__init__()
        self.setStyleSheet("""
            QCheckBox
            {
                background-color:"""+config.backgroundColor+""";
                color:"""+config.accentColor+""";
            }
            QCheckBox::indicator
            {
                width: 50px;
                height: 50px;
            }
            QCheckBox::indicator:unchecked
            {
                image: url(images/unchecked.png);
            }
            QCheckBox::indicator:checked
            {
                image: url(images/checked.png);
            }
            
                                    """)
        self.setText(text)

        infoTextFont = QFont()
        infoTextFont.setFamily("Serif")
        infoTextFont.setFixedPitch( True )
        infoTextFont.setPointSize( parent.width() / 90 )
        self.setFont(infoTextFont)
        self.setMouseTracking(True)

    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.PointingHandCursor)

class DayButton(QPushButton):
    def __init__(self, parent, text, buttonArr):
        super(DayButton, self).__init__()
        # store the text
        self.buttonText = text
        self.buttonArr = buttonArr
        # set the size
        self.setFixedHeight(parent.height() / 30)
        self.setFixedWidth(parent.width() / 30)
        # add the day of the monht to it
        self.setText(text)
        # set the stylesheet
        self.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: """ + config.backgroundColor + """;
            background-color: """ + config.accentColor + """;
        """)
        monthButtonFont = QFont()
        monthButtonFont.setFamily("Serif")
        monthButtonFont.setFixedPitch( True )
        monthButtonFont.setPointSize( parent.height() / 55 )
        self.setFont(monthButtonFont)   
        self.clicked.connect(self.buttonClicked)
        # each button has a TOGGLE variable
        self.toggle = False

    def buttonClicked(self):
        global toggledDays
        # if you click it and it is currently untoggled
        if self.toggle == False:
            # make sure we untoggle the other day that is currently toggled
            for i in range(0, len(self.buttonArr)):
                # if we find a toggled button we untoggle it
                if self.buttonArr[i].toggle == True:
                    self.buttonArr[i].buttonClicked()
            
            # then toggle the current button
            self.setStyleSheet("""
                text-align:center;
                border-radius: 5px;
                color: """ + config.accentColor + """;
                background-color: """ + config.numberColor + """;
            """)
            self.toggle = True

        else:
            self.setStyleSheet("""
                text-align:center;
                border-radius: 5px;
                color: """ + config.backgroundColor + """;
                background-color: """ + config.accentColor + """;
            """)
            self.toggle = False


class NameScreen(QWidget):
    def __init__(self, parent):
        super(NameScreen, self).__init__()
        global nameScreen
        config.nameScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(20)
        self.nameText = QLabel(self)
        self.nameText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        nameTextFont = QFont()
        nameTextFont.setFamily("Serif")
        nameTextFont.setFixedPitch( True )
        nameTextFont.setPointSize( parent.width() / 50 )
        self.nameText.setFont(nameTextFont)
        self.nameText.setText('Select a day and its matching Zoom schedule:')     
        self.nameText.setAlignment(QtCore.Qt.AlignCenter)

        self.infoText = QLabel(self)
        self.infoText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.numberColor + """;
        """)
        infoTextFont = QFont()
        infoTextFont.setFamily("Serif")
        infoTextFont.setFixedPitch( True )
        infoTextFont.setPointSize( parent.width() / 110 )
        self.infoText.setFont(infoTextFont)
        self.infoText.setText('(Pick the schedule that matches that day even if you did not get attendance for all of the sessions)')
        self.infoText.setAlignment(QtCore.Qt.AlignCenter)

        # create 5 horizontal layouts to store the 31 buttons :(
        self.horizontalLayouts = []
        for i in range(0, 6):
            self.horizontalLayouts.append(QHBoxLayout())
            self.horizontalLayouts[i].setSpacing(10)
            if i < 5:
                self.horizontalLayouts[i].addStretch(-1)

        # create the buttons
        self.buttonArr = []
        for i in range(1, 32):
            self.buttonArr.append(DayButton(parent, str(i), self.buttonArr))

        curButton = 0
        # add the buttons to the horizontal layouts
        for i in range(0, 6):
            # for each horizontal layout we will add 6 buttons
            for j in range(0, 6):
                if curButton + j <= 30:
                    self.horizontalLayouts[i].addWidget(self.buttonArr[curButton + j])
            self.horizontalLayouts[i].addStretch(-1)
            curButton += 6

        # create a vertical layout to store the horizontal layouts
        self.daysVert = QVBoxLayout()
        self.daysVert.setSpacing(10)

        # add the day rows to this
        for i in range(0, 6):
            self.daysVert.addLayout(self.horizontalLayouts[i])

        # create hor layout to hor center scenario dropdown
        self.dropLayout = QHBoxLayout()
        self.dropLayout.setSpacing(60)
        self.dropLayout.addStretch(-1)
        
        # add the days to this
        self.dropLayout.addLayout(self.daysVert)

        # create a dropdown with the different scenarios
        # create the dropdown menu
        self.scenario = QComboBox()
        # add the dropdown to the hor layout
        self.dropLayout.addWidget(self.scenario)
        # add a stretch after
        self.dropLayout.addStretch(-1)
        # set the size of the month select dropdown menu
        self.scenario.setFixedHeight(parent.height() / 10)
        self.scenario.setFixedWidth(parent.width() / 3)
        # add the months to the combobox
        self.scenario.addItem('Select your remote schedule:')
        self.scenario.addItem('1.\n\tTwo Zoom-only Sessions\n\tTwo Hybrid Zoom Sessions')
        self.scenario.addItem('2.\n\tTwo Zoom-only Sessions\n\tOne Hybrid Zoom Session')
        self.scenario.addItem('3.\n\tOne Zoom-only Session\n\tTwo Hybrid Zoom Sessions')
        self.scenario.addItem('4.\n\tOne Zoom-only Session\n\tOne Hybrid Zoom Session')
        self.scenario.addItem('5. Two Zoom-only Sessions')
        self.scenario.addItem('6. One Zoom-only Session')
        self.scenario.addItem('7. Two Hybrid Zoom Sessions')
        self.scenario.addItem('8. One Hybrid Zoom Session')
        
        self.scenario.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: black;
            background-color: white;
            selection-background-color:"""+config.accentColor+""";
        """)
        scenarioFont = QFont()
        scenarioFont.setFamily("Serif")
        scenarioFont.setFixedPitch( True )
        scenarioFont.setPointSize( parent.width() / 80 )
        self.scenario.setFont(scenarioFont) 

        # create the checkbox for review
        self.reviewCheck = Check(parent, "\tReview (zoom-only or hybrid)")
        
        # create horizontal layout for the button
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.addStretch(-1)
        self.buttonLayout.addWidget(self.reviewCheck)
        self.buttonLayout.addStretch(-1)
        
        # create a button to continue
        self.nameButton = QPushButton()
        # create the function for when its pressed
        self.nameButton.clicked.connect(self.pressedContinue)
        self.nameButton.setText("Continue")
        self.nameButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.nameButton.setFixedHeight(parent.width() / 15)
        self.nameButton.setFixedWidth(parent.width() / 6)

        # create a back button
        self.backButton = QPushButton()
        # create the function for when its pressed
        self.backButton.clicked.connect(self.pressedBack)
        self.backButton.setText("Back")
        self.backButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.backButton.setFixedHeight(parent.width() / 15)
        self.backButton.setFixedWidth(parent.width() / 6)

        # set the font of teh button
        nameButtonFont = QFont()
        nameButtonFont.setFamily("Serif")
        nameButtonFont.setFixedPitch( True )
        nameButtonFont.setPointSize( parent.width() / 44 )
        self.nameButton.setFont(nameButtonFont)
        self.backButton.setFont(nameButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(10)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.backButton)
        self.hLayout.addWidget(self.nameButton)
        self.hLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.nameText)
        self.vLayout.addWidget(self.infoText)

        self.vLayout.addLayout(self.dropLayout)
        self.vLayout.addLayout(self.buttonLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)

        self.setMouseTracking(True)

    def pressedContinue(self):
        # when the start button is pressed it sends you to the month selection screen
        #config.stack.setCurrentIndex(10)

        # turn the dropdown red if they have not selected a scenario or review
        isToggled = False
        for i in range(0, 31):
            if self.buttonArr[i].toggle == True:
                isToggled = True
        # if they have not selected a day yet then flash the days
        if isToggled == False or self.scenario.currentIndex() == 0:
            for i in range(0, config.flashNumber):
                if isToggled == False:
                    for j in range(0, 31):
                        self.buttonArr[j].setStyleSheet("""
                            text-align:center;
                            border-radius: 5px;
                            color: """ + config.accentColor + """;
                            background-color: """ + config.numberColor + """;
                        """)
                if self.scenario.currentIndex() == 0:
                    self.scenario.setStyleSheet("""
                        text-align:center;
                        border-radius: 5px;
                        color: black;
                        background-color:"""+config.numberColor+""";
                        selection-background-color:"""+config.accentColor+""";
                    """)
                QtTest.QTest.qWait(config.waitTime)
                if isToggled == False:
                    for j in range(0, 31):
                        self.buttonArr[j].setStyleSheet("""
                            text-align:center;
                            border-radius: 5px;
                            color: """ + config.backgroundColor + """;
                            background-color: """ + config.accentColor + """;
                        """)
                if self.scenario.currentIndex() == 0:
                    self.scenario.setStyleSheet("""
                            text-align:center;
                            border-radius: 5px;
                            color: black;
                            background-color: white;
                            selection-background-color:"""+config.accentColor+""";
                        """)
                QtTest.QTest.qWait(config.waitTime)
        # if they gave all the info needed
        else:
            config.stack.setCurrentIndex(10)
            # get the day from the buttonArr
            day = ''
            for i in range(0, 31):
                if self.buttonArr[i].toggle == True:
                    day = self.buttonArr[i].text()
            # if they had a review
            if self.reviewCheck.isChecked() == True:
                # scenario 1
                if self.scenario.currentIndex() == 1:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2\nHybrid Session 1: {}-1h\nHybrid Session 2: {}-2h\nReview: {}r".format(day, day, day, day, day))
                elif self.scenario.currentIndex() == 2:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2\nHybrid Session: {}h\nReview: {}r".format(day, day, day, day))
                elif self.scenario.currentIndex() == 3:
                    config.helpScreen.helpText.setText("Zoom Session: {}\nHybrid Session 1: {}-1h\nHybrid Session 2: {}-2h\nReview: {}r".format(day, day, day, day))
                elif self.scenario.currentIndex() == 4:
                    config.helpScreen.helpText.setText("Zoom Session: {}\nHybrid Session: {}h\nReview: {}r".format(day, day, day))
                elif self.scenario.currentIndex() == 5:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2\nReview: {}r".format(day, day, day))
                elif self.scenario.currentIndex() == 6:
                    config.helpScreen.helpText.setText("Zoom Session: {}\nReview: {}r".format(day, day))
                elif self.scenario.currentIndex() == 7:
                    config.helpScreen.helpText.setText("Hybrid Session 1: {}-1h\nHybrid Session 2: {}-2h\nReview: {}r".format(day, day, day))
                elif self.scenario.currentIndex() == 8:
                    config.helpScreen.helpText.setText("Hybrid Session: {}h\nReview: {}r".format(day, day))
            else:
                    # scenario 1
                if self.scenario.currentIndex() == 1:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2\nHybrid Session 1: {}-1h\nHybrid Session 2: {}-2h".format(day, day, day, day))
                elif self.scenario.currentIndex() == 2:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2\nHybrid Session: {}h".format(day, day, day))
                elif self.scenario.currentIndex() == 3:
                    config.helpScreen.helpText.setText("Zoom Session: {}\nHybrid Session 1: {}-1h\nHybrid Session 2: {}-2h".format(day, day, day))
                elif self.scenario.currentIndex() == 4:
                    config.helpScreen.helpText.setText("Zoom Session: {}\nHybrid Session: {}h".format(day, day))
                elif self.scenario.currentIndex() == 5:
                    config.helpScreen.helpText.setText("Zoom Session 1: {}-1\nZoom Session 2: {}-2".format(day, day))
                elif self.scenario.currentIndex() == 6:
                    config.helpScreen.helpText.setText("Zoom Session: {}".format(day))
                elif self.scenario.currentIndex() == 7:
                    config.helpScreen.helpText.setText("Hybrid Session 1: {}-1h\nHybrid Session 2: {}-2h".format(day, day))
                elif self.scenario.currentIndex() == 8:
                    config.helpScreen.helpText.setText("Hybrid Session: {}h".format(day))

    def pressedBack(self):
        config.stack.setCurrentIndex(0)
        
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
