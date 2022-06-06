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

class HelpScreen(QWidget):
    def __init__(self, parent):
        super(HelpScreen, self).__init__()
        global helpScreen
        config.helpScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(0)
        self.helpText = QLabel(self)
        self.helpText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        helpTextFont = QFont()
        helpTextFont.setFamily("Serif")
        helpTextFont.setFixedPitch( True )
        helpTextFont.setPointSize( parent.width() / 50 )
        self.helpText.setFont(helpTextFont)
        self.helpText.setAlignment(QtCore.Qt.AlignCenter)
        
        # create a button to continue
        self.helpButton = QPushButton()
        # create the function for when its pressed
        self.helpButton.clicked.connect(self.pressedStart)
        self.helpButton.setText("Back")
        self.helpButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.helpButton.setFixedHeight(parent.width() / 12)
        self.helpButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        helpButtonFont = QFont()
        helpButtonFont.setFamily("Serif")
        helpButtonFont.setFixedPitch( True )
        helpButtonFont.setPointSize( parent.width() / 35 )
        self.helpButton.setFont(helpButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.helpButton)
        self.hLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.helpText)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)

        self.setMouseTracking(True)

    def pressedStart(self):
        # when the start button is pressed it sends you to the month selection screen
        config.stack.setCurrentIndex(9)
        
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
