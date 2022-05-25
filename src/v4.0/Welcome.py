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
from PyQt5.QtWidgets import QPushButton, QDesktopWidget, QSizePolicy
from PyQt5.QtWidgets import QVBoxLayout, QScrollBar
from PyQt5.QtWidgets import QWidget, QFrame
from PyQt5.QtCore import Qt, QRect, QSize, QRectF
from PyQt5.QtWidgets import QWidget, QPlainTextEdit, QTextEdit
from PyQt5.QtGui import QColor, QPainter, QTextFormat, QLinearGradient
import os
import ctypes

import TitleBar, config, AutoGrid, Welcome, FirstWindow, Month

class WelcomeScreen(QWidget):
    def __init__(self, parent):
        super(WelcomeScreen, self).__init__()
        global welcomeScreen
        config.welcomeScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(40)
        self.welcomeText = QLabel(self)
        self.welcomeText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        font = QFont()
        font.setFamily("Serif")
        font.setFixedPitch( True )
        font.setPointSize( parent.width() / 20 )
        self.welcomeText.setFont(font)
        self.welcomeText.setText('Welcome to <b style="color: #A3BE8C;">AutoGrid!</b>')
        self.welcomeText.setAlignment(QtCore.Qt.AlignCenter)

        self.pressStart = QLabel(self)
        self.pressStart.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        font2 = QFont()
        font2.setFamily("Serif")
        font2.setFixedPitch( True )
        font2.setPointSize( parent.width() / 45 )
        self.pressStart.setFont(font2)
        #self.pressStart.setText('Click <a style="color: #A3BE8C;">Start</a>:')
        self.pressStart.setAlignment(QtCore.Qt.AlignCenter)
        
        # create a button to continue
        self.welcomeButton = QPushButton()
        # create the function for when its pressed
        self.welcomeButton.clicked.connect(self.pressedStart)
        self.welcomeButton.setText("Start")
        self.welcomeButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.welcomeButton.setFixedHeight(parent.width() / 12)
        self.welcomeButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        font3 = QFont()
        font3.setFamily("Serif")
        font3.setFixedPitch( True )
        font3.setPointSize( parent.width() / 35 )
        self.welcomeButton.setFont(font3)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.welcomeButton)
        self.hLayout.addStretch(-1)

        # create horizontal layout for the help button
        self.helpLayout = QHBoxLayout()
        self.helpLayout.setSpacing(20)
        self.helpLayout.addStretch(-1)

        # add qlabel asking if they need help naming attendance sheets
        self.helpText = QLabel(self)
        self.helpText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor2 + """;
        """)
        helpFont = QFont()
        helpFont.setFamily("Serif")
        helpFont.setFixedPitch( True )
        helpFont.setPointSize( parent.width() / 100 )
        self.helpText.setFont(helpFont)
        self.helpText.setText('How to name Zoom attendance reports: ')
        self.helpText.setAlignment(QtCore.Qt.AlignCenter)

        # add the text to the helpLayout
        self.helpLayout.addWidget(self.helpText)

        # create help button
        self.helpButton = QPushButton()
        # create the function for when its pressed
        self.helpButton.clicked.connect(self.pressedGuide)
        self.helpButton.setText("Guide")
        self.helpButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.helpButton.setFixedHeight(parent.width() / 40)
        self.helpButton.setFixedWidth(parent.width() / 12)
        
        # set the font of teh button
        helpButtonFont2 = QFont()
        helpButtonFont2.setFamily("Serif")
        helpButtonFont2.setFixedPitch( True )
        helpButtonFont2.setPointSize( parent.width() / 80 )
        self.helpButton.setFont(helpButtonFont2)

        # add button to the help layout
        self.helpLayout.addWidget(self.helpButton)
        self.helpLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.welcomeText)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        self.vLayout.addStretch(-1)
        self.vLayout.addLayout(self.helpLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)

        self.setMouseTracking(True)

    def pressedStart(self):
        # when the start button is pressed it sends you to the month selection screen
        config.stack.setCurrentIndex(1)
    
    def pressedGuide(self):
        # when the start button is pressed it sends you to the name help
        config.stack.setCurrentIndex(9)
        
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
