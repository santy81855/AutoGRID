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

class ZoomScreen(QWidget):
    def __init__(self, parent):
        super(ZoomScreen, self).__init__()
        global zoomScreen
        config.zoomScreen = self
        
        # create the layout
        self.vLayout = QVBoxLayout()
        self.vLayout.setSpacing(80)
        self.zoomText = QLabel(self)
        self.zoomText.setStyleSheet("""
            border: none;
            vertical-align: top;
            text-align:center;
            color: """ + config.accentColor + """;
        """)
        zoomFont = QFont()
        zoomFont.setFamily("Serif")
        zoomFont.setFixedPitch( True )
        zoomFont.setPointSize( parent.width() / 32 )
        self.zoomText.setFont(zoomFont)
        self.zoomText.setText("Upload your Zoom attendance reports:")
        self.zoomText.setAlignment(QtCore.Qt.AlignCenter)

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
        self.numFiles = QLabel("0 Attendance Reports")
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
        numFilesFont.setPointSize( parent.width() / 53 )
        self.numFiles.setFont(numFilesFont)    

        # add the number of files qlabel to the horizontal layout
        self.fileLayout.addWidget(self.numFiles)
        # add the browse button to the horizontal layout
        self.fileLayout.addWidget(self.fileSelect)
        # add another stretch
        self.fileLayout.addStretch(-1) 
        
        # create a button to continue
        self.zoomButton = QPushButton()
        # create the function for when its pressed
        self.zoomButton.clicked.connect(self.pressedContinue)
        self.zoomButton.setText("Continue")
        self.zoomButton.setStyleSheet("""
            border-radius: 20px;
            background-color: """ + config.accentColor + """;
            color: """ + config.backgroundColor + """;
                                        """)
        # set the size of the button
        self.zoomButton.setFixedHeight(parent.width() / 12)
        self.zoomButton.setFixedWidth(parent.width() / 3)
        
        # set the font of teh button
        zoomButtonFont = QFont()
        zoomButtonFont.setFamily("Serif")
        zoomButtonFont.setFixedPitch( True )
        zoomButtonFont.setPointSize( parent.width() / 35 )
        self.zoomButton.setFont(zoomButtonFont)

        # add the button to a horizontal layout
        self.hLayout = QHBoxLayout()
        self.hLayout.setSpacing(0)
        self.hLayout.addStretch(-1)
        self.hLayout.addWidget(self.zoomButton)
        self.hLayout.addStretch(-1)

        # add a stretch before the button for extra separation
        self.vLayout.addStretch(-1)
        # add the qlabel to the vlayout
        self.vLayout.addWidget(self.zoomText)
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
        global zoom_attendance_reports
        global zoom_days
        # get the names of all the attendance reports
        # allow both xlsx and csv so that people can't mess up
        aTuple = QFileDialog.getOpenFileNames(self, 'open files', '', 'CSV files (*.csv)')
        # place those names in a global list if they selected the correct number of files
        if len(aTuple[0]) > 0:
            AutoGrid.zoom_attendance_reports *= 0
            AutoGrid.zoom_days *= 0
            for i in range(0, len(aTuple[0])):
                AutoGrid.zoom_attendance_reports.append(aTuple[0][i])
            # now get the day of each file and put it in another list with the same indexes
            for i in range(0, len(aTuple[0])):
                temp = AutoGrid.zoom_attendance_reports[i]
                lastIndex = len(temp) - 1
                # if there is a dash and an h then it is extra long
                if ('-1' in temp or '-2' in temp) and 'h' in temp:
                    # if it is a double digit number we have to go back 9 spaces
                    # this means it is a double digit date
                    if lastIndex - 8 >= 0 and temp[lastIndex - 8].isdigit():
                        AutoGrid.zoom_days.append(temp[lastIndex - 8] + temp[lastIndex - 7] + temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])
                    # this means it is a single digit date
                    else:
                        AutoGrid.zoom_days.append(temp[lastIndex - 7] + temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])
                    
                # if there is a dash at all it means it is a double day
                elif ('-1' in temp) or ('-2' in temp):
                    # this means it is a double digit date
                    if lastIndex - 7 >= 0 and temp[lastIndex - 7].isdigit():
                        AutoGrid.zoom_days.append(temp[lastIndex - 7] + temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])
                    # this means it is a single digit date
                    else:
                        AutoGrid.zoom_days.append(temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])

                # this means it is a review day
                elif temp[lastIndex - 4] == 'r':
                    # we have 2 options
                    # double digit day with an r at the end
                    if (lastIndex - 6) >= 0 and temp[lastIndex - 6].isdigit():
                        AutoGrid.zoom_days.append(temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])
                    # single digit day with an r at the end
                    else:
                        AutoGrid.zoom_days.append(temp[lastIndex - 5] + temp[lastIndex - 4])
                # this means it is a hybrid session
                elif temp[lastIndex - 4] == 'h':
                    # we have 2 options
                    # double digit day with an h at the end
                    if (lastIndex - 6) >= 0 and temp[lastIndex - 6].isdigit():
                        AutoGrid.zoom_days.append(temp[lastIndex - 6] + temp[lastIndex - 5] + temp[lastIndex - 4])
                    # single digit day with an h at the end
                    else:
                        AutoGrid.zoom_days.append(temp[lastIndex - 5] + temp[lastIndex - 4])
                # this means it is a double digit normal day
                elif (lastIndex - 5) >= 0 and temp[lastIndex - 5].isdigit():
                    AutoGrid.zoom_days.append(temp[lastIndex - 5] + temp[lastIndex - 4])
                # this means that it is a single digit normal date
                else:
                    AutoGrid.zoom_days.append(temp[lastIndex - 4])
            config.zoomScreen.numFiles.setText(str(len(aTuple[0])) + ' Attendance Reports')
            print(AutoGrid.zoom_days)

    def pressedContinue(self):
        config.stack.setCurrentIndex(3)
            
        
        # convert it into a month
        #config.stack.setCurrentIndex(2)
    
    def mouseMoveEvent(self, event):
        QApplication.setOverrideCursor(Qt.ArrowCursor)
        
        
