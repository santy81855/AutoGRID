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

class PersonScreen(QWidget):
    def __init__(self, parent):
        super(PersonScreen, self).__init__()
        global personScreen
        config.personScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(80)
        self.personText = QLabel(self)
        self.personText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        gridFont = QFont()
        gridFont.setFamily("Serif")
        gridFont.setFixedPitch( True )
        gridFont.setPointSize( parent.width() / 32 )
        self.personText.setFont(gridFont)
        self.personText.setText("Upload Your In-Person Attendance:")
        self.personText.setAlignment(QtCore.Qt.AlignCenter)

        # create a horizontal layout to put the file dialogue in
        self.fileLayout = QHBoxLayout()
        # create some spacing between the numfiles and the browse button
        self.fileLayout.setSpacing(20)
        self.fileLayout.addStretch(-1)

        # create the browse button
        self.fileSelect = QPushButton('Browse')
        self.fileSelect.clicked.connect(self.pressedBrowse)
        # set the size of the browse button
        self.fileSelect.setFixedHeight(parent.width() / 22)
        self.fileSelect.setFixedWidth(parent.width() / 6)
        #border: none;
        #vertical-align: top;
        #text-align:center;
        self.fileSelect.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: """ + config.backgroundColor + """;
            background-color: """ + config.accentColor + """;
        """)
        browseButtonFont = QFont()
        browseButtonFont.setFamily("Serif")
        browseButtonFont.setFixedPitch( True )
        browseButtonFont.setPointSize( parent.width() / 40 )
        self.fileSelect.setFont(browseButtonFont)    

        # create a label that will display the number of files that were uploaded to show the user
        self.numFiles = QLabel()
        self.numFiles.setFixedHeight(parent.width() / 24)
        self.numFiles.setFixedWidth(parent.width() / 3)
        self.numFiles.setStyleSheet("""
            background-color: white;
            color: black; 
            text-align:center;
                                    """)
        self.numFiles.setAlignment(QtCore.Qt.AlignCenter)
        numFilesFont = QFont()
        numFilesFont.setFamily("Serif")
        numFilesFont.setFixedPitch( True )
        numFilesFont.setPointSize( parent.width() / 90 )
        self.numFiles.setFont(numFilesFont)    

        # add the number of files qlabel to the horizontal layout
        self.fileLayout.addWidget(self.numFiles)
        # add the browse button to the horizontal layout
        self.fileLayout.addWidget(self.fileSelect)
        # add another stretch
        self.fileLayout.addStretch(-1) 
        
        # create a button to continue
        self.personButton = QPushButton()
        # create the function for when its pressed
        self.personButton.clicked.connect(self.pressedContinue)
        self.personButton.setText("Continue")
        self.personButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.personButton.setFixedHeight(parent.width() / 12)
        self.personButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        personButtonFont = QFont()
        personButtonFont.setFamily("Serif")
        personButtonFont.setFixedPitch( True )
        personButtonFont.setPointSize( parent.width() / 35 )
        self.personButton.setFont(personButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.personButton)
        self.hLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.personText)
        # add the horizontal layout of the dropdown
        self.vLayout.addLayout(self.fileLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)
        # track mouse movement so we can change cursor
        self.setMouseTracking(True)

    def pressedBrowse(self):
        fileName=QFileDialog.getOpenFileName(self, 'open file', '', 'XLSX files (*.xlsx)')
        global attendance_sheet_name
        AutoGrid.attendance_sheet_name = fileName[0]


        if '/' in str(fileName[0]):
            shortName = str(fileName[0]).split('/')
            config.personScreen.numFiles.setText(shortName[len(shortName) - 1].replace('.xlsx', ''))
        else:
            config.personScreen.numFiles.setText(fileName[0])
        

    def pressedContinue(self):
        global inPerson
        # if the qlabel is blank then there is no in person attendance
        if self.numFiles.text() == '':
            AutoGrid.inPerson == False
        else:
            AutoGrid.inPerson == True
            
        config.stack.setCurrentIndex(5)
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
