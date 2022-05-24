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

class ObserverNumber(QLabel):
    def __init__(self, parent, text):
        super(ObserverNumber, self).__init__()
        self.setText(text)
        self.setFixedHeight(parent.width() / 24)
        self.setFixedWidth(parent.width() / 3)
        self.setStyleSheet("""
            background-color:""" + config.backgroundColor + """;
            color:""" + config.accentColor + """; 
            text-align:center;
                                    """)
        self.setAlignment(QtCore.Qt.AlignCenter)
        observerNumberFont = QFont()
        observerNumberFont.setFamily("Serif")
        observerNumberFont.setFixedPitch( True )
        observerNumberFont.setPointSize( parent.width() / 53 )
        self.setFont(observerNumberFont)   

class ObserverName(QPlainTextEdit):
    def __init__(self, parent, text):
        super(ObserverName, self).__init__()
        self.setPlaceholderText(text)
        self.setFixedHeight(parent.width() / 24)
        self.setFixedWidth(parent.width() / 5)
        self.setStyleSheet("""
            background-color:white;
            color:""" + config.numberColor + """; 
                                    """)
        textFont = QFont()
        textFont.setFamily("Serif")
        textFont.setFixedPitch( True )
        textFont.setPointSize( parent.width() / 53 )
        self.setFont(textFont)    
        self.setMouseTracking(True)
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.IBeamCursor)

class ObservationScreen(QWidget):
    def __init__(self, parent):
        super(ObservationScreen, self).__init__()
        global observationScreen
        config.observationScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(60)
        self.observedText = QLabel(self)
        self.observedText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        observedFont = QFont()
        observedFont.setFamily("Serif")
        observedFont.setFixedPitch( True )
        observedFont.setPointSize( parent.width() / 44 )
        self.observedText.setFont(observedFont)
        self.observedText.setText("Enter the name of each mentor/GA that observed you:")
        self.observedText.setAlignment(QtCore.Qt.AlignCenter)

        # create a horizontal layout to put the label and textbox in
        self.mentor1Layout = QHBoxLayout()
        self.mentor2Layout = QHBoxLayout()
        self.mentor3Layout = QHBoxLayout()
        self.mentor4Layout = QHBoxLayout()
        # create some spacing between the label and the textbox
        self.mentor1Layout.setSpacing(20)
        self.mentor1Layout.addStretch(-1)
        
        self.mentor2Layout.setSpacing(20)
        self.mentor2Layout.addStretch(-1)

        self.mentor3Layout.setSpacing(20)
        self.mentor3Layout.addStretch(-1)

        self.mentor4Layout.setSpacing(20)
        self.mentor4Layout.addStretch(-1)

        # create a label that will display the order of the mentors
        self.observer1 = ObserverNumber(parent, "Observer #1")
        self.observer2 = ObserverNumber(parent, "Observer #2")
        self.observer3 = ObserverNumber(parent, "Observer #3")
        self.observer4 = ObserverNumber(parent, "Observer #4")

        # create a text box that will allow them to enter the names
        self.name1 = ObserverName(parent, "First")
        self.name2 = ObserverName(parent, "First")
        self.name3 = ObserverName(parent, "First")
        self.name4 = ObserverName(parent, "First")

        self.last1 = ObserverName(parent, "Last")
        self.last2 = ObserverName(parent, "Last")
        self.last3 = ObserverName(parent, "Last")
        self.last4 = ObserverName(parent, "Last")

        # add the label to the mentors layouts
        self.mentor1Layout.addWidget(self.observer1)
        self.mentor2Layout.addWidget(self.observer2)
        self.mentor3Layout.addWidget(self.observer3)
        self.mentor4Layout.addWidget(self.observer4)
        #
        # add the text boxes to each horizontal layout
        self.mentor1Layout.addWidget(self.name1)
        self.mentor1Layout.addWidget(self.last1)
        self.mentor2Layout.addWidget(self.name2)
        self.mentor2Layout.addWidget(self.last2)
        self.mentor3Layout.addWidget(self.name3)
        self.mentor3Layout.addWidget(self.last3)
        self.mentor4Layout.addWidget(self.name4)
        self.mentor4Layout.addWidget(self.last4)
        # add another stretch to each horizontal layout
        self.mentor1Layout.addStretch(-1) 
        self.mentor2Layout.addStretch(-1) 
        self.mentor3Layout.addStretch(-1) 
        self.mentor4Layout.addStretch(-1) 

        # add the title label to the vlayout
        self.vLayout.addWidget(self.observedText)
        '''
        # now depending on the number of observations we will display them
        if AutoGrid.num_observations >= 1:
            self.vLayout.addLayout(self.mentor1Layout)
        if AutoGrid.num_observations >= 2:
            self.vLayout.addLayout(self.mentor2Layout)
        if AutoGrid.num_observations >= 3:
            self.vLayout.addLayout(self.mentor3Layout)
        if AutoGrid.num_observations >= 4:
            self.vLayout.addLayout(self.mentor4Layout)
        '''
        # create a button to continue
        self.observationButton = QPushButton()
        # create the function for when its pressed
        self.observationButton.clicked.connect(self.pressedContinue)
        self.observationButton.setText("Continue")
        self.observationButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.observationButton.setFixedHeight(parent.width() / 12)
        self.observationButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        observationButtonFont = QFont()
        observationButtonFont.setFamily("Serif")
        observationButtonFont.setFixedPitch( True )
        observationButtonFont.setPointSize( parent.width() / 35 )
        self.observationButton.setFont(observationButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.observationButton)
        self.hLayout.addStretch(-1)
        '''
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch after the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)
        '''
        # track mouse movement so we can change cursor
        self.setMouseTracking(True)

    def pressedContinue(self):
        global mentors
        # if they try to continue without typing all the names
        # then set any empty text boxes to have a red background
        empty = 0
        if (self.name1.toPlainText() == '' or self.last1.toPlainText() == '') and AutoGrid.num_observations >= 1:
            empty = 1
            self.name1.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
            self.last1.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
        else:
            self.name1.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
            self.last1.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
        if (self.name2.toPlainText() == '' or self.last2.toPlainText() == '') and AutoGrid.num_observations >= 2:
            empty = 2
            self.name2.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
            self.last2.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
        else:
            self.name2.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
            self.last2.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
        if (self.name3.toPlainText() == '' or self.last3.toPlainText() == '') and AutoGrid.num_observations >= 3:
            empty = 3
            self.name3.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
            self.last3.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
        else:
            self.name3.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
            self.last3.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
        if (self.name4.toPlainText() == '' or self.last4.toPlainText() == '') and AutoGrid.num_observations >= 4:
            empty = 4           
            self.name4.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
            self.last4.setStyleSheet("""
            background-color:""" + config.numberColor + """;
            color:white; 
                                    """)
        else:
            self.name4.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
            self.last4.setStyleSheet("""
            background-color: white;
            color:""" + config.numberColor + """; 
                                    """)
    
        # if none of the names are emepty then just store all variables
        if empty == 0:        
            # add the names to the mentor's dictionary
            #AutoGrid.mentors[firstName] = lastName
            if AutoGrid.num_observations >= 1:
                AutoGrid.mentors[str(self.name1.toPlainText())] = str(self.last1.toPlainText())
            if AutoGrid.num_observations >= 2:
                AutoGrid.mentors[str(self.name2.toPlainText())] = str(self.last2.toPlainText())
            if AutoGrid.num_observations >= 3:
                AutoGrid.mentors[str(self.name3.toPlainText())] = str(self.last3.toPlainText())
            if AutoGrid.num_observations >= 4:
                AutoGrid.mentors[self.name4.toPlainText()] = str(self.last4.toPlainText())
            
            # If we have any exam days then go to that screen next
            if AutoGrid.num_exams > 0:
                config.stack.setCurrentIndex(7)       
            # otherwise just run the program and launch loading screen
            else:
                config.stack.setCurrentIndex(8)       
                AutoGrid.runAutoGrid()
                AutoGrid.runAutoGridZoom()
                config.loadingScreen.loadingText.setText("Your grid is complete!")
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
