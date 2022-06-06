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

class ExtraScreen(QWidget):
    def __init__(self, parent):
        super(ExtraScreen, self).__init__()
        global extraScreen
        config.extraScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(30)
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
        observedFont.setPointSize( parent.width() / 40 )
        self.observedText.setFont(observedFont)
        self.observedText.setText("How many times were you observed over Zoom?")
        self.observedText.setAlignment(QtCore.Qt.AlignCenter)

        # create a horizontal layout to put the dropdown in
        self.dropdownLayout = QHBoxLayout()
        self.dropdownLayout.setSpacing(0)
        self.dropdownLayout.addStretch(-1)

        # create the dropdown menu
        self.observedSelect = QComboBox()
        # set the size of the observed select dropdown menu
        self.observedSelect.setFixedHeight(parent.width() / 22)
        self.observedSelect.setFixedWidth(parent.width() / 15)
        # add the months to the combobox
        self.observedSelect.addItem('0')
        self.observedSelect.addItem('1')
        self.observedSelect.addItem('2')
        self.observedSelect.addItem('3')
        self.observedSelect.addItem('4')

        # add signal for if it gets changed
        self.observedSelect.currentIndexChanged.connect(self.changed)
        
        #border: none;
        #vertical-align: top;
        #text-align:center;
        self.observedSelect.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: black;
            background-color: white;
        """)
        observedDropdownFont = QFont()
        observedDropdownFont.setFamily("Serif")
        observedDropdownFont.setFixedPitch( True )
        observedDropdownFont.setPointSize( parent.width() / 40 )
        self.observedSelect.setFont(observedDropdownFont)    

        # add the dropdown to the horizontal layout
        self.dropdownLayout.addWidget(self.observedSelect)
        # add another stretch
        self.dropdownLayout.addStretch(-1) 
        
        # create a button to continue
        self.extraButton = QPushButton()
        # create the function for when its pressed
        self.extraButton.clicked.connect(self.pressedContinue)
        self.extraButton.setText("Run Program")
        self.extraButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.extraButton.setFixedHeight(parent.width() / 12)
        self.extraButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        extraButtonFont = QFont()
        extraButtonFont.setFamily("Serif")
        extraButtonFont.setFixedPitch( True )
        extraButtonFont.setPointSize( parent.width() / 35 )
        self.extraButton.setFont(extraButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.extraButton)
        self.hLayout.addStretch(-1)
        ######################################
        ######################################
        ######################################
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
        examFont.setPointSize( parent.width() / 45 )
        self.examText.setFont(examFont)
        self.examText.setText("How many exams did the class have this month?")
        self.examText.setAlignment(QtCore.Qt.AlignCenter)

        # create a horizontal layout to put the dropdown in
        self.dropdownLayout2 = QHBoxLayout()
        self.dropdownLayout2.setSpacing(0)
        self.dropdownLayout2.addStretch(-1)

        # create the dropdown menu
        self.examSelect = QComboBox()
        # set the size of the observed select dropdown menu
        self.examSelect.setFixedHeight(parent.width() / 22)
        self.examSelect.setFixedWidth(parent.width() / 15)
        # add the months to the combobox
        self.examSelect.addItem('0')
        self.examSelect.addItem('1')
        self.examSelect.addItem('2')
        self.examSelect.addItem('3')
        self.examSelect.addItem('4')

        # add signal for if it gets changed
        self.examSelect.currentIndexChanged.connect(self.changed)
        
        #border: none;
        #vertical-align: top;
        #text-align:center;
        self.examSelect.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: black;
            background-color: white;
        """)
        examDropdownFont = QFont()
        examDropdownFont.setFamily("Serif")
        examDropdownFont.setFixedPitch( True )
        examDropdownFont.setPointSize( parent.width() / 40 )
        self.examSelect.setFont(examDropdownFont)    

        # add the dropdown to the horizontal layout
        self.dropdownLayout2.addWidget(self.examSelect)
        # add another stretch
        self.dropdownLayout2.addStretch(-1) 
        ######################
        ###################
        #################

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.observedText)
        # add the horizontal layout of the dropdown
        self.vLayout.addLayout(self.dropdownLayout)
        # add the qlabel for the exams
        self.vLayout.addWidget(self.examText)
        # add the horizontal layout of the dropdown
        self.vLayout.addLayout(self.dropdownLayout2)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)
        # track mouse movement so we can change cursor
        self.setMouseTracking(True)

    def changed(self):
        # if the user changes from the default 0 then change the 'Run Program' button to say continue
        if self.observedSelect.currentIndex() != 0 or self.examSelect.currentIndex() != 0:
            self.extraButton.setText("Continue")
        else:
            self.extraButton.setText("Run Program")

    def pressedContinue(self):
        # when the continue button is pressed it sends you to the next screen if a month is selected
        global num_observations
        global num_exams
        # get the current index and set it as the number of observations
        AutoGrid.num_observations = self.observedSelect.currentIndex()
        AutoGrid.num_exams = self.examSelect.currentIndex()

        # make the title of the exam screen customized
        if AutoGrid.num_exams == 1:
            config.examScreen.examText.setText("Select the day of the exam:")
        else:
            config.examScreen.examText.setText("Select the " + str(AutoGrid.num_exams) + " exam days:")

        # if the user has any observations we need to send them to the observation screen
        # which is an index higher
        if self.observedSelect.currentIndex() > 0:
            # We have to update the observation page here before calling it
            if AutoGrid.num_observations >= 1:
                config.observationScreen.vLayout.addLayout(config.observationScreen.mentor1Layout)
            if AutoGrid.num_observations >= 2:
                config.observationScreen.vLayout.addLayout(config.observationScreen.mentor2Layout)
            if AutoGrid.num_observations >= 3:
                config.observationScreen.vLayout.addLayout(config.observationScreen.mentor3Layout)
            if AutoGrid.num_observations >= 4:
                config.observationScreen.vLayout.addLayout(config.observationScreen.mentor4Layout)
            # add a stretch before the button for extra separation
            config.observationScreen.vLayout.addStretch(-1)
            # add the horizontal layout
            config.observationScreen.vLayout.addLayout(config.observationScreen.hLayout)
            # add a stretch after the button for extra separation
            config.observationScreen.vLayout.addStretch(-1)
            config.observationScreen.setLayout(config.observationScreen.vLayout)
            
            # move to the observation page
            config.stack.setCurrentIndex(6)
            
        # if the exams are more than 0 send them to the exam screen
        elif self.examSelect.currentIndex() > 0:
            config.stack.setCurrentIndex(7)
        # otherwise just run the program and launch the loading screen
        else:
            config.stack.setCurrentIndex(8)
            AutoGrid.runAutoGrid()
            AutoGrid.runAutoGridZoom()
            config.loadingScreen.loadingText.setText("Your grid is complete!")
            
        

        # move to the next screenn (zoom)
        #config.stack.setCurrentIndex(2)
            
        
        # convert it into a month
        #config.stack.setCurrentIndex(2)
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
