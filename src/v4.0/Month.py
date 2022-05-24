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

class MonthScreen(QWidget):
    def __init__(self, parent):
        super(MonthScreen, self).__init__()
        global monthScreen
        config.monthScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(80)
        self.monthText = QLabel(self)
        self.monthText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        font = QFont()
        font.setFamily("Serif")
        font.setFixedPitch( True )
        font.setPointSize( parent.width() / 30 )
        self.monthText.setFont(font)
        self.monthText.setText("Select the month of attendance:")
        self.monthText.setAlignment(QtCore.Qt.AlignCenter)

        # create a horizontal layout to put the dropdown in
        self.dropdownLayout = QHBoxLayout()
        self.dropdownLayout.setSpacing(0)
        self.dropdownLayout.addStretch(-1)

        # create the dropdown menu
        self.monthSelect = QComboBox()
        # set the size of the month select dropdown menu
        self.monthSelect.setFixedHeight(parent.width() / 22)
        self.monthSelect.setFixedWidth(parent.width() / 3)
        # add the months to the combobox
        self.monthSelect.addItem('Select a month:')
        self.monthSelect.addItem('January')
        self.monthSelect.addItem('February')
        self.monthSelect.addItem('March')
        self.monthSelect.addItem('April')
        self.monthSelect.addItem('May')
        self.monthSelect.addItem('June')
        self.monthSelect.addItem('July')
        self.monthSelect.addItem('August')
        self.monthSelect.addItem('September')
        self.monthSelect.addItem('October')
        self.monthSelect.addItem('November')
        self.monthSelect.addItem('December')
        #border: none;
        #vertical-align: top;
        #text-align:center;
        self.monthSelect.setStyleSheet("""
            text-align:center;
            border-radius: 5px;
            color: """ + config.backgroundColor + """;
            background-color: """ + config.accentColor + """;
        """)
        font2 = QFont()
        font2.setFamily("Serif")
        font2.setFixedPitch( True )
        font2.setPointSize( parent.width() / 40 )
        self.monthSelect.setFont(font2)    

        # add the dropdown to the horizontal layout
        self.dropdownLayout.addWidget(self.monthSelect)
        # add another stretch
        self.dropdownLayout.addStretch(-1) 
        
        # create a button to continue
        self.monthButton = QPushButton()
        # create the function for when its pressed
        self.monthButton.clicked.connect(self.pressedContinue)
        self.monthButton.setText("Continue")
        self.monthButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.monthButton.setFixedHeight(parent.width() / 12)
        self.monthButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        font3 = QFont()
        font3.setFamily("Serif")
        font3.setFixedPitch( True )
        font3.setPointSize( parent.width() / 35 )
        self.monthButton.setFont(font3)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.monthButton)
        self.hLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.monthText)
        # add the horizontal layout of the dropdown
        self.vLayout.addLayout(self.dropdownLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the horizontal layout
        self.vLayout.addLayout(self.hLayout)
        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        self.setLayout(self.vLayout)
        # track mouse movement so we can change cursor
        self.setMouseTracking(True)

    def pressedContinue(self):
        # when the continue button is pressed it sends you to the next screen if a month is selected
        global current_month
        # get the current index and set it as the current month
        if self.monthSelect.currentIndex() != 0:
            AutoGrid.current_month = self.monthSelect.currentIndex()
            # move to the next screenn (zoom)
            config.stack.setCurrentIndex(2)
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
