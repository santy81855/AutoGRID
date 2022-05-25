import sys
# to get the working monitor size
from win32api import GetMonitorInfo, MonitorFromPoint
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtGui import QCursor, QMouseEvent, QFont, QKeySequence, QSyntaxHighlighter, QTextCharFormat, QBrush, QTextCursor
from PyQt5.QtCore import QPoint, pyqtSignal, QRegExp
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect, QEasingCurve
from PyQt5.QtCore import QObject, QMimeData
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QCompleter, QFileDialog, QGraphicsDropShadowEffect, QComboBox
from PyQt5.QtWidgets import QHBoxLayout, QTextEdit, QPlainTextEdit, QShortcut, QScrollArea
from PyQt5.QtWidgets import QLabel, QStackedWidget, QMessageBox
from PyQt5.QtWidgets import QPushButton, QDesktopWidget, QSizePolicy
from PyQt5.QtWidgets import QVBoxLayout, QScrollBar
from PyQt5.QtWidgets import QWidget, QFrame
from PyQt5.QtCore import Qt, QRect, QSize, QRectF
from PyQt5.QtWidgets import QWidget, QPlainTextEdit, QTextEdit
from PyQt5.QtGui import QColor, QPainter, QTextFormat, QLinearGradient
import os
import ctypes

import TitleBar, config, AutoGrid, Welcome, FirstWindow

class DayButton(QPushButton):
    def __init__(self, parent, text):
        super(DayButton, self).__init__()
        # store the text
        self.buttonText = text
        # set the size
        self.setFixedHeight(parent.height() / 18)
        self.setFixedWidth(parent.width() / 18)
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
        monthButtonFont.setPointSize( parent.width() / 40 )
        self.setFont(monthButtonFont)   
        self.clicked.connect(self.buttonClicked)
        # each button has a TOGGLE variable
        self.toggle = False

    def buttonClicked(self):
        # if you click it and it is currently untoggled
        if self.toggle == False:
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
        

class ExamScreen(QWidget):
    def __init__(self, parent):
        super(ExamScreen, self).__init__()
        global examScreen
        config.examScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(20)
        self.examText = QLabel(self)
        self.examText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        examFont = QFont()
        examFont.setFamily("Serif")
        examFont.setFixedPitch( True )
        examFont.setPointSize( parent.height() / 20 )
        self.examText.setFont(examFont)
        # if they only had one exam 
        if AutoGrid.num_exams == 1:
            self.examText.setText("Select the day of the exam:")
        else:
            self.examText.setText("Select the " + str(AutoGrid.num_exams) + " exam days:")
        self.examText.setAlignment(QtCore.Qt.AlignCenter)

        # create 5 horizontal layouts to store the 31 buttons :(
        self.horizontalLayouts = []
        for i in range(0, 6):
            self.horizontalLayouts.append(QHBoxLayout())
            self.horizontalLayouts[i].setSpacing(20)
            self.horizontalLayouts[i].addStretch(-1)

        # create the buttons
        self.buttonArr = []
        for i in range(1, 32):
            self.buttonArr.append(DayButton(parent, str(i)))

        curButton = 0
        # add the buttons to the horizontal layouts
        for i in range(0, 6):
            # for each horizontal layout we will add 6 buttons
            for j in range(0, 6):
                if curButton + j <= 30:
                    self.horizontalLayouts[i].addWidget(self.buttonArr[curButton + j])
            self.horizontalLayouts[i].addStretch(-1)
            curButton += 6
        
        # add a stretch before the title for extra separation
        self.vLayout.addStretch(-1)
        # add the title to the vlayout
        self.vLayout.addWidget(self.examText)
        # add a stretch after the title
        self.vLayout.addStretch(-1)
        # now add the horizontal layouts to the vertical layout
        for i in range(0, 6):
            self.vLayout.addLayout(self.horizontalLayouts[i])
        # add a stretch
        self.vLayout.addStretch(-1)
        
        # create a button to continue
        self.examButton = QPushButton()
        # create the function for when its pressed
        self.examButton.clicked.connect(self.pressedContinue)
        self.examButton.setText("Continue")
        self.examButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.examButton.setFixedHeight(parent.width() / 12)
        self.examButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        examButtonFont = QFont()
        examButtonFont.setFamily("Serif")
        examButtonFont.setFixedPitch( True )
        examButtonFont.setPointSize( parent.width() / 35 )
        self.examButton.setFont(examButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.examButton)
        self.hLayout.addStretch(-1)

        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)
        # track mouse movement so we can change cursor
        self.setMouseTracking(True)

    def pressedContinue(self):
        # once they press continue we will check which buttons are toggled and store their info
        # AutoGrid.exam_days.append(int(day))
        global exam_days
        for i in range(0, 31):
            if self.buttonArr[i].toggle == True:
                AutoGrid.exam_days.append(int(self.buttonArr[i].text()))
        
        # finally we run the program, so we go to the loading screen
        config.stack.setCurrentIndex(8)
        AutoGrid.runAutoGrid()
        AutoGrid.runAutoGridZoom()
        config.loadingScreen.loadingText.setText("Your grid is complete!")
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
