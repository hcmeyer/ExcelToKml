#!/usr/bin/env python
#
#   basic imports
#
import  os
import  platform
import  sys
#
#   gui toolkit
#
from    PyQt4.QtCore    import *
from    PyQt4.QtGui     import *
from    PyQt4.QtWebKit  import *
#
#   libraries
#
#   geocoding
#
from    geopy   import  geocoders
#
#   basic xml
#
from lxml import etree
#
#   Excel files
#
from xlrd         import open_workbook
from xlutils.copy import copy
#
#   Kml
#
from pykml.factory import KML_ElementMaker as KML
from pykml.factory import ATOM_ElementMaker as ATOM
from pykml.factory import GX_ElementMaker as GX
#
#   aux dialog to display help
#
import  helpform
#
#   compiled resources
#
import  excelToKmlQrc
#
try:
    _fromUtf8 = QString.fromUtf8
except AttributeError:
    _fromUtf8 = lambda s: s

__version__ = "1.0.0"

class   MainWindow( QMainWindow):
    nameColumnIndex        = 0
    descriptionColumnIndex = 1
    streetColumnIndex      = 2
    townColumnIndex        = 3
    stateColumnIndex       = 4
    zipColumnIndex         = 5
    telephoneColumnIndex   = 6
    websiteColumnIndex     = 7
    latitudeColumnIndex    = 8
    longitudeColumnIndex   = 9
#
#   put all the size numbers here for tinkering
#
    appWidth        = 1000
    appHeight       = 450
#
    appMinWidth     = ((2*appWidth)/3)
    appMinHeight    = ((2*appHeight)/3)
    appMaxWidth     = ((3*appWidth)/2)
    appMaxHeight    = ((3*appHeight)/2)
#    
    topLeftWidth    = ((11*appWidth)/20)
    topRightWidth   = (appWidth-topLeftWidth)
    topHeight       = ((8*appHeight)/9)
    bottomHeight    = (appHeight-topHeight)
#       
    bluePushpin    = [ "BluePushpin",
                        ":/Icon/blue-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png']
    greenPushpin    = [ "GreenPushpin",
                        ":/Icon/grn-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png']
    ltBluPushpin    = [ "LightBluePushpin",
                        ":/Icon/ltblu-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/ltblu-pushpin.png']
    pinkPushpin     = [ "PinkPushpin",
                        ":/Icon/pink-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/pink-pushpin.png']
    purplePushpin   = [ "PurplePushpin",
                        ":/Icon/purple-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/purple-pushpin.png']
    redPushpin      = [ "RedPushpin",
                        ":/Icon/red-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png']
    whtPushpin      = [ "WhitePushpin",
                        ":/Icon/wht-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png']
    yellowPushpin   = [ "YellowPushpin",
                        ":/Icon/ylw-pushpin.png",
                        'http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png']

    kmlFolderIcons  = [ bluePushpin, greenPushpin,  ltBluPushpin, pinkPushpin,
                        purplePushpin, redPushpin,  whtPushpin,   yellowPushpin]

    def __init__(   self, parent = None):

        super(  MainWindow, self).__init__( parent)
        #
        #   data state flags
        #
        self.colsValid      = False
        self.SheetsDirty    = False
        self.treeDirty      = False
        self.rowDirty       = False
        self.geocodingValid = False
        #
        #   file names
        #
        self.excelInputFileName  = None
        self.excelOutputFileName = None
        self.kmlOutputFileName   = None
        self.baseFileName  = None
        self.dirName  = None
        self.baseName = None
        #
        #   data objects
        #
        self.readWorkBook  = None
        self.writeWorkBook = None
        #    
        self.sheets     = None
        self.sheetNames = None
        self.sheetIconIndexes = None
        self.sheetIndex = -1
        self.rowIndex   = -1
        self.kmlDoc     = None
        #
        #   link to web service
        #
        self.googleGeoCoder = geocoders.Google()
        #
        # gui
        #
        self.setupUi()
        #
        #   status bar at bottom of main window
        #
        status  = self.statusBar()
        status.setSizeGripEnabled(  False)
        status.showMessage( "Ready", 5000)
        #
        #   file actions, excel flavor, and kml flavors
        #
        excelFileOpenAction    = self.createAction( "&Open...",           self.excelFileOpen,
                                                    QKeySequence.Open,   "Icon/excelFileOpen",
                                                    "Open an existing Excel file")
        fileQuitAction         = self.createAction( "&Quit",              self.close,    
                                                    "Ctrl+Q",             "Icon/exit",  
                                                    "Close the application")
        self.kmlFileSaveAction = self.createAction( "&Save",              self.kmlFileSave, 
                                                    QKeySequence.Save,    "Icon/kmlFileSave", 
                                                    "Save the kml file")
        self.excelFileSaveAction = self.createAction(  "&Save",              self.excelFileSave, 
                                                  icon="Icon/excelFileSave", 
                                                  tip="Save the Excel file")
        self.excelFileSaveAsAction = self.createAction("Save &As...",   self.excelFileSaveAs, 
                                                  icon="Icon/excelFileSaveAs", 
                                                  tip="Save the Excel file using a new name")
        self.kmlFileSaveAsAction = self.createAction(  "Save &As...",        self.kmlFileSaveAs, 
                                                  icon="Icon/kmlFileSaveAs", 
                                                  tip="Save the kml file using a new name")
        #                                        
        #   data actions, all kml flavor
        # 
        self.dataKmlAction   = self.createAction( "&Kml",    self.dataKml,   "Ctrl+K",
                                                  "Icon/kml",     "Convert sheet data to Kml")
        self.dataGeoAction   = self.createAction( "&Geo",    self.dataGeo,   "Ctrl+G",
                                                  "Icon/geocode", "Check sheet data for Geocoding")
        #
        #   help actions
        #
        helpAboutAction = self.createAction( "&About",     self.helpAbout,  "Ctrl+F1",
                                             "helpAbout", "Display Help - About")
        helpHelpAction  = self.createAction( "&Help",      self.helpHelp,  "F1",
                                             "helpHelp",  "Display Help")
        #
        #   create menus
        #
        self.excelMenu = self.menuBar().addMenu( "&ExcelFile")
        self.kmlMenu   = self.menuBar().addMenu( "&KmlFile")
        self.helpMenu  = self.menuBar().addMenu( "&Help")
        #
        #   connect actions to menus
        #
        self.excelMenuActions    = ( excelFileOpenAction,      None,
                                     self.excelFileSaveAction, self.excelFileSaveAsAction, 
                                     None,                     fileQuitAction)
        #
        #   file actions are added to file menu in updateExcelMenu
        #   to implement lru list
        #
        self.kmlMenuActions    = ( self.dataKmlAction,     self.dataGeoAction,
                                   None,                     
                                   self.kmlFileSaveAction, self.kmlFileSaveAsAction)
        self.addActions(    self.kmlMenu,
                            self.kmlMenuActions)
        #
        self.helpMenuActions   = ( helpAboutAction, helpHelpAction)
        self.addActions(    self.helpMenu,
                            self.helpMenuActions)
        #
        #   build tool bars
        #
        excelToolBar = self.addToolBar(  "Excel")
        kmlToolBar   = self.addToolBar(  "Kml")
        #
        #   connect actions to tool bars
        #
        excelToolBar.setObjectName(  "ExcelToolBar")
        self.addActions(    excelToolBar,
                            (excelFileOpenAction, 
                            self.excelFileSaveAction, 
                            self.excelFileSaveAsAction))
        kmlToolBar.setObjectName(  "KmlToolBar")
        self.addActions(    kmlToolBar,
                            (self.dataKmlAction,    self.dataGeoAction, 
                            self.kmlFileSaveAction, self.kmlFileSaveAsAction))
        #
        #   connect signals and slots
        #
        self.wireUpEvents()
        #                   
        #   settings saved and restored between invocations
        #
        settings    = QSettings()
        #
        self.recentFiles    = settings.value(   "RecentFiles").toStringList()
        #
        size = settings.value(   "MainWindow/Size",
               QVariant( QSize( self.appWidth, self.appHeight))).toSize()
        self.resize(    size)
        #
        position    = settings.value(   "MainWindows/Position",
                        QVariant( QPoint( 0, 0))).toPoint()
        self.move(  position)
        #
        self.restoreState(  settings.value( "MainWindow/State").toByteArray())
        #
        self.vertSplitter.restoreState(  settings.value( "VertSplitter/State").toByteArray())
        self.horizSplitter.restoreState( settings.value( "HorizSplitter/State").toByteArray())
        #
        #
        self.setWindowTitle( "kml Viewer [*]")
        #
        self.updateExcelMenu()
        self.updateExcelActions()
        self.updateKmlActions()
        #
        self.updateSheetUi()
        self.updateRowUi()
        #
        QTimer.singleShot(  0, self.loadInitialFile)
        #
        return
        
    def setupUi( self):
        #
        #
        #
        self.setWindowIcon( QIcon(":/Icon/kml.png"))
        #
        #       vert splitter will become central widget
        #
        self.vertSplitter = QSplitter( )
        self.vertSplitter.setOrientation( Qt.Vertical)
        self.vertSplitter.setMinimumSize( self.appMinWidth, 
                                          self.appMinHeight)
        self.vertSplitter.setMaximumSize( self.appMaxWidth,
                                          self.appMaxHeight)
        self.vertSplitter.setObjectName( "vertSplitter")
        #
        #       horiz spliiter is top part of vert splitter
        #                        
        self.horizSplitter = QSplitter( self.vertSplitter)
        self.horizSplitter.setOrientation( Qt.Horizontal)
        self.horizSplitter.setObjectName( "horizSplitter")
        #
        #       spreadsheet is edit controls and labels for spreadsheet data
        #       will be left part of horiz spliiter, top part of vert splitter
        #
        self.spreadsheetScrollArea = QScrollArea( self.horizSplitter)
        self.spreadsheetScrollArea.setObjectName( "spreadsheetScrollArea")
        #
        self.spreadsheetScrollAreaContents = QWidget()
        self.spreadsheetScrollAreaContents.setFixedSize( self.topLeftWidth, 
                                                         self.topHeight)
        self.spreadsheetScrollAreaContents.setObjectName( "spreadsheetScrollAreaContents")
        #
        self.spreadsheetScrollArea.setWidget( self.spreadsheetScrollAreaContents)        
        #
        #       tree widget shows structure and data of kml
        #       right part of horiz splitter, top part of vert splitter
        #
        self.treeWidget = QTreeWidget( self.horizSplitter)
        self.treeWidget.setObjectName( "treeView")
        self.treeWidget.clear()
        self.treeWidget.setColumnCount( 2)
        self.treeWidget.setHeaderLabels( [ "Keys", "Text"])
        #       left first        
        self.horizSplitter.addWidget( self.spreadsheetScrollArea)
        #       right second
        self.horizSplitter.addWidget( self.treeWidget)
        #
        #       logListWidget shows all status messages for session
        #       bottom part of vert splitter
        #        
        self.logListWidget = QListWidget( self.horizSplitter)
        self.logListWidget.setMinimumHeight( 20)
        self.logListWidget.setObjectName( "logListWidget")
        #       top first
        self.vertSplitter.addWidget( self.horizSplitter)
        #       bottom second
        self.vertSplitter.addWidget( self.logListWidget)               
        #
        self.setCentralWidget( self.vertSplitter)
        #
        #       these sizes will be reset by save/restore function        
        #
        self.vertSplitter.setSizes(  [self.topHeight,
                                      self.bottomHeight])
        self.horizSplitter.setSizes( [self.topLeftWidth,
                                      self.topRightWidth])
        # 
        #       composite layout for top left, spreadsheet data 
        #       outermost is vert box
        #
        self.verticalLayout = QVBoxLayout( self.spreadsheetScrollAreaContents)
        self.verticalLayout.setContentsMargins( 3, 0, 3, 0)
        self.verticalLayout.setObjectName( "verticalSpreadsheetLayout")
        #
        #       these labels appear above the grid of labels
        #       and edit controls for spreadsheet data
        #
        #       file name data will be two lines high,
        #       with a line break between dir name and file name
        #
        self.fileNameLabel = QLabel()
        self.fileNameLabel.setText( "File: Name")
        self.fileNameLabel.setObjectName( "fileNameLabel")
        self.fileNameLabel.setMinimumHeight( 20)
        self.fileNameLabel.setWordWrap( True)
        #
        self.verticalLayout.addWidget( self.fileNameLabel)
        #        
        self.sheetNameLabel = QLabel()
        self.sheetNameLabel.setText( "Sheet: Name")
        self.sheetNameLabel.setObjectName( "sheetNameLabel")
        #
        self.verticalLayout.addWidget( self.sheetNameLabel)
        #
        #       inside of vert box layout is 
        #       grid layout with 5 cols
        #       one row, addresses, has 6 cols in 
        #       an embedded horiz box layout
        #        
        self.gridSpreadsheetLayout = QGridLayout( )
        self.gridSpreadsheetLayout.setObjectName( "gridSpreadsheetLayout")
        #
        self.verticalLayout.addLayout( self.gridSpreadsheetLayout)
        #
        #       first row of grid is sheet count and controls
        #        
        #
        self.sheetControlLayout = QHBoxLayout( )
        #        
        self.prevSheetButton = QPushButton()
        self.prevSheetButton.setText( "Prev Sheet")
        self.prevSheetButton.setObjectName( "prevSheetButton")
        self.prevSheetButton.setIcon( QIcon(":/Icon/back.png"))
        self.prevSheetButton.setToolTip( "Goto Previous Worksheet")
        self.prevSheetButton.setEnabled( False)
        self.sheetControlLayout.addWidget( self.prevSheetButton)
        #self.gridSpreadsheetLayout.addWidget( self.prevSheetButton, 0, 0, 1, 1)
        #
        self.sheetNofMlabel = QLabel()
        self.sheetNofMlabel.setText( "Sheet N of M")
        self.sheetNofMlabel.setObjectName( "sheetNofMlabel")
        self.sheetControlLayout.addWidget( self.sheetNofMlabel)
        #self.gridSpreadsheetLayout.addWidget( self.sheetNofMlabel, 0, 2, 1, 1)
        #
        self.sheetIconDropdown  = QComboBox()
        for kfi in self.kmlFolderIcons:
            self.sheetIconDropdown.addItem( QIcon(kfi[1]), QString( kfi[0]))    
        self.sheetIconDropdown.setObjectName( "sheetIconDropdown")
        self.sheetIconDropdown.setEnabled( False)
        self.sheetControlLayout.addWidget( self.sheetIconDropdown)
        #self.gridSpreadsheetLayout.addWidget( self.sheetIconDropdown, 0, 3, 1, 1)
        #
        self.nextSheetButton = QPushButton()
        self.nextSheetButton.setText( "Next Sheet")
        self.nextSheetButton.setObjectName( "nextSheetButton")
        self.nextSheetButton.setIcon( QIcon(":/Icon/forward.png"))
        self.nextSheetButton.setToolTip( "Goto Next Worksheet")
        self.nextSheetButton.setEnabled( False)
        self.sheetControlLayout.addWidget( self.nextSheetButton)
        #self.gridSpreadsheetLayout.addWidget( self.nextSheetButton, 0, 4, 1, 1)
        #
        self.gridSpreadsheetLayout.addLayout( self.sheetControlLayout, 0, 0, 1, 5)
        #
        #       spreadsheet row data
        #
        self.nameLabel = QLabel()
        self.nameLabel.setText( "Name:")
        self.nameLabel.setObjectName( "nameLabel")
        self.gridSpreadsheetLayout.addWidget( self.nameLabel, 1, 0, 1, 1)
        #
        self.nameEdit = QLineEdit()
        self.nameEdit.setObjectName( "nameEdit")
        self.gridSpreadsheetLayout.addWidget( self.nameEdit, 1, 1, 1, 4)
        #
        self.descriptionLabel = QLabel()
        self.descriptionLabel.setText( "Description:")
        self.descriptionLabel.setObjectName( "descriptionLabel")
        self.gridSpreadsheetLayout.addWidget( self.descriptionLabel, 2, 0, 1, 1)
        #
        self.descriptionTextEdit = QPlainTextEdit()
        self.descriptionTextEdit.setMaximumSize( QSize(16777215, 60))
        self.descriptionTextEdit.setObjectName( "descriptionTextEdit")
        self.gridSpreadsheetLayout.addWidget( self.descriptionTextEdit, 2, 1, 2, 4)
        #
        self.streetLabel = QLabel()
        self.streetLabel.setText( "Street:")
        self.streetLabel.setObjectName( "streetLabel")
        self.gridSpreadsheetLayout.addWidget( self.streetLabel, 4, 0, 1, 1)
        #
        self.streetEdit = QLineEdit()
        self.streetEdit.setObjectName( "streetEdit")
        self.gridSpreadsheetLayout.addWidget( self.streetEdit, 4, 1, 1, 4)
        #
        #--------------------------------------------------------------------
        #       address row, 3 spreadsheet cols
        #       this layout row actually has 6 cols,
        #       in an embedded Horiz box layout
        #
        self.townStateZipLayout = QHBoxLayout( )
        #        
        self.townLabel = QLabel()
        self.townLabel.setText( "Town:")
        self.townLabel.setObjectName( "townLabel")
        self.townStateZipLayout.addWidget( self.townLabel)
        #
        self.townEdit = QLineEdit()
        self.townEdit.setObjectName( "townEdit")
        self.townStateZipLayout.addWidget( self.townEdit)
        #
        self.stateLabel = QLabel()
        self.stateLabel.setText( "State:")
        self.stateLabel.setObjectName( "stateLabel")
        self.townStateZipLayout.addWidget( self.stateLabel)
        #
        self.stateEdit = QLineEdit()
        self.stateEdit.setObjectName( "stateEdit")
        self.townStateZipLayout.addWidget( self.stateEdit)
        #
        self.zipLabel = QLabel()
        self.zipLabel.setText( "Zip:")
        self.zipLabel.setObjectName( "zipLabel")
        self.townStateZipLayout.addWidget( self.zipLabel)
        #
        self.zipEdit = QLineEdit()
        self.zipEdit.setObjectName( "zipEdit")
        self.townStateZipLayout.addWidget( self.zipEdit)
        #
        self.gridSpreadsheetLayout.addLayout( self.townStateZipLayout, 5, 0, 1, 5)
        #
        #-------------------------------------------------------------        
        #
        self.telephoneLabel = QLabel()
        self.telephoneLabel.setText( "Telephone:")
        self.telephoneLabel.setObjectName( "telephoneLabel")
        self.gridSpreadsheetLayout.addWidget( self.telephoneLabel, 6, 0, 1, 1)
        #
        self.telephoneEdit = QLineEdit()
        self.telephoneEdit.setObjectName( "telephoneEdit")
        self.gridSpreadsheetLayout.addWidget( self.telephoneEdit, 6, 1, 1, 4)
        #
        self.websiteLabel = QLabel()
        self.websiteLabel.setText( "Website")
        self.websiteLabel.setObjectName( "websiteLabel")
        self.gridSpreadsheetLayout.addWidget( self.websiteLabel, 7, 0, 1, 1)
        #
        self.websiteEdit = QLineEdit()
        self.websiteEdit.setObjectName( "websiteEdit")
        self.gridSpreadsheetLayout.addWidget( self.websiteEdit, 7, 1, 1, 4)
        #
        self.latitudeLabel = QLabel()
        self.latitudeLabel.setText( "Latitude:")
        self.latitudeLabel.setObjectName( "latitudeLabel")
        self.gridSpreadsheetLayout.addWidget( self.latitudeLabel, 8, 0, 1, 1)
        #
        self.latitudeEdit = QLineEdit()
        self.latitudeEdit.setObjectName( "latitudeEdit")
        self.gridSpreadsheetLayout.addWidget( self.latitudeEdit, 8, 1, 1, 2)
        #
        self.lookupGeoCodeButton = QPushButton()
        self.lookupGeoCodeButton.setText( "Lookup")
        self.lookupGeoCodeButton.setObjectName( "lookupGeoCodeButton")
        self.lookupGeoCodeButton.setIcon( QIcon(":/Icon/geocode.png"))
        self.lookupGeoCodeButton.setToolTip( "Lookup latitude, longitude for this row")
        self.lookupGeoCodeButton.setEnabled( False)
        self.gridSpreadsheetLayout.addWidget( self.lookupGeoCodeButton, 8, 3, 2, 1)
        #
        self.longitudeLabel = QLabel()
        self.longitudeLabel.setText( "Longitude:")
        self.longitudeLabel.setObjectName( "longitudeLabel")
        self.gridSpreadsheetLayout.addWidget( self.longitudeLabel, 9, 0, 1, 1)
        #
        self.longitudeEdit = QLineEdit()
        self.longitudeEdit.setObjectName( "longitudeEdit")
        self.gridSpreadsheetLayout.addWidget( self.longitudeEdit, 9, 1, 1, 2)
        #
        #       last row is row count and controls
        #        
        self.prevRowButton = QPushButton()
        self.prevRowButton.setText( "Prev Row")
        self.prevRowButton.setObjectName( "prevRowButton")
        self.prevRowButton.setIcon( QIcon(":/Icon/back.png"))
        self.prevRowButton.setToolTip( "Goto Previous Row")
        self.prevRowButton.setEnabled( False)
        self.gridSpreadsheetLayout.addWidget( self.prevRowButton, 10, 0, 1, 1)
        #
        self.undoRowButton = QPushButton()
        self.undoRowButton.setText( "Undo")
        self.undoRowButton.setObjectName( "undoRowButton")
        self.undoRowButton.setIcon( QIcon(":/Icon/undo.png"))
        self.undoRowButton.setToolTip( "Undo changes to this row")
        self.undoRowButton.setEnabled( False)
        self.gridSpreadsheetLayout.addWidget( self.undoRowButton, 10, 1, 1, 1)
        #
        self.rowNofMlabel = QLabel()
        self.rowNofMlabel.setText( "Row N of M")
        self.rowNofMlabel.setObjectName( "rowNofMlabel")
        self.gridSpreadsheetLayout.addWidget( self.rowNofMlabel, 10, 2, 1, 1)
        #
        self.updateRowButton = QPushButton()
        self.updateRowButton.setText( "Update")
        self.updateRowButton.setObjectName( "updateRowButton")
        self.updateRowButton.setToolTip( "Update this row with changes")
        self.updateRowButton.setIcon( QIcon(":/Icon/update.png"))
        self.updateRowButton.setEnabled( False)
        self.gridSpreadsheetLayout.addWidget( self.updateRowButton, 10, 3, 1, 1)
        #
        self.nextRowButton = QPushButton()
        self.nextRowButton.setText( "Next Row")
        self.nextRowButton.setObjectName( "nextRowButton")
        self.nextRowButton.setIcon( QIcon(":/Icon/forward.png"))
        self.nextRowButton.setToolTip( "Goto next row")
        self.nextRowButton.setEnabled( False)
        self.gridSpreadsheetLayout.addWidget( self.nextRowButton, 10, 4, 1, 1)
        #        
        return
        
    def addActions( self, target, actions):
        for action in actions:
            if action is None:
                target.addSeparator()
            else:
                target.addAction(   action)
        return
        
    def addRecentFile(  self,   fname):
        if fname is None:
            return
        if not self.recentFiles.contains( fname):
            self.recentFiles.prepend( QString( fname))
            while self.recentFiles.count() > 9:
                self.recentFiles.takeLast()
        return
    #       
    #   edit controls changed
    #    
    def changedNameData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedDescriptionData( self, newModified):
        #
        if newModified:
            self.setRowDirty( True)
        #
        return
        
    def changedStreetData( self, newText):
        #
        self.setRowDirty( True)
        #
        return
        
    def changedTownData( self, newText):
        #
        self.setRowDirty( True)
        #
        return
        
    def changedStateData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedZipData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedTelephoneData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedWebsiteData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedLatitudeData( self, newText):
        #
        self.setRowDirty( True)
        #
        return

    def changedLongitudeData( self, newText):
        #
        self.setRowDirty( True)
        #
        return
        
    def closeEvent( self, event):
        if not self.okToContinue():
            event.ignore
            return

        settings    = QSettings()

        fileName    = QVariant( QString( self.excelInputFileName)) \
                        if self.excelInputFileName is not None else QVariant()
        settings.setValue(  "LastFile", fileName)

        recentFiles = QVariant( self.recentFiles) \
                        if self.recentFiles else QVariant()
        settings.setValue( "RecentFiles",           recentFiles)

        settings.setValue( "MainWindow/Size",       QVariant( self.size( )))

        settings.setValue( "MainWindows/Position",  QVariant( self.pos()))
        
        settings.setValue( "MainWindow/State",      QVariant( self.saveState()))
        #
        settings.setValue( "VertSplitter/State" ,   QVariant( self.vertSplitter.saveState()))
        settings.setValue( "HorizSplitter/State" ,  QVariant( self.horizSplitter.saveState()))
        #
        return                       
    #    
    #   convience helper
    #
    def createAction(   self,   text,   slot = None,    shortCut = None,    
                        icon = None,    tip = None,     checkable = False,  
                        signal = "triggered()"):
        action  = QAction(  text, self)
        if icon is not None:
            action.setIcon( QIcon(  ":/%s.png" % icon))
        if shortCut is not None:
            action.setShortcut( shortCut)
        if tip is not None:
            action.setToolTip(  tip)
            action.setStatusTip(tip)
        if slot is not None:
            self.connect(   action, SIGNAL( signal), slot)
        if checkable:
            action.setCheckable( True)
        return action
    
    def dataGeo(    self):
        #
        #   check all rows of all sheets for
        #   valid latitude, longitude
        #
        self.validateGeoCoding()
        self.updateSheetUi()
        self.updateRowUi()
        #
        return

    def dataKml( self):
        #
        #   produce the kml tree doc
        #
        #   got data ?
        #
        if( self.sheetNames is None
        or  self.sheets is None):
            message = "no Excel data"
            self.updateStatus( message)
            return
        #    
        self.kmlDoc = self.kmlDocFromSheets( self.sheetNames,
                                             self.sheetIconIndexes,
                                             self.sheets)
        if self.kmlDoc is None:
            message = "Kml data invalid"
            self.updateStatus( message)
            return;
        #
        self.populateTreeWidget( self.kmlDoc)
        self.setTreeDirty( True)
        #    
        message = "Kml tree built"
        self.updateStatus( message)
        #
        return

###################################################################################################
#
#   excel file functions
# 
    def excelFileOpen(   self):
        if not self.okToContinue():
            return
        
        dirName = os.path.dirname(  self.excelInputFileName) \
                    if self.excelInputFileName is not None else "."
        fname   = unicode(  QFileDialog.getOpenFileName( self,
                            "Excel Files - Choose Workbook",    
                             dirName, "Excel files (*.xls *.XLS)"))
        if fname:
            self.excelLoadFile( fname)
        return
        
    def excelFileSave(  self):
        if( self.sheets is None
        or  self.writeWorkBook is None):
            self.updateStatus(  "No Excel data to Save")
            return
            
        if self.excelOutputFileName is None:
            self.excelFileSaveAs()
            return
        #    
        if self.excelWriteData():
            self.setSheetsDirty( False)
            self.updateStatus(  "Excel data saved as %s" % self.excelOutputFileName)
            return
        #
        self.updateStatus(  "Failed to save Excel data file %s" % self.excelOutputFileName)
        #
        return

    def excelFileSaveAs(  self):
        if( self.sheets is None
        or  self.writeWorkBook is None):
            self.updateStatus(  "No Excel Data to Save")
            return
        dirName = ( os.path.dirname( self.excelOutputFileName)
                    if self.excelOutputFileName is not None else ".")
        fName   = unicode(  QFileDialog.getSaveFileName( self,
                            "Excel Files - Save Output",    
                             dirName, "Excel files (*.xls *.XLS)"))
        if not fName:
            return
        if "." not in fName:
            fName   += ".xls"
        #
        self.excelOutputFileName   = fName 
        self.excelFileSave()
        
        return
        
    def excelLoadFile(   self, fName=None):
        #
        #
        #
        if fName is None:
            action  = self.sender()
            if isinstance(  action, QAction):
                fName   = unicode(  action.data().toString())
                if not self.okToContinue():
                    return
            else:
                return
        #
        if fName:
            self.excelInputFileName   = None
            #
            #   workbook is the excel data from the .xls file
            #   see xlrd library documentation for details
            #
            self.readWorkBook = open_workbook(fName)
            #
            self.writeWorkBook = None
            #
            #   sheets is a python 3d list of strings
            #   with data from the workbook
            #
            #   sheets[] are the sheets of the workbook
            #   sheets[][] are the rows of the sheet
            #   sheets[][][] are the columns of the row
            #
            self.sheets     = []
            self.sheetNames = []
            self.sheetIconIndexes = []
            self.sheetIndex = -1
            self.rowIndex   = -1
            self.kmlDoc     = None
            self.populateTreeWidget( None)
                    
            for s in self.readWorkBook.sheets():
                #   print
                #   print 'Sheet:',s.name
                #   print
                self.sheetNames.append( s.name)
                self.sheetIconIndexes.append( 0)
                rows = []
                self.sheetIndex += 1
                for row in range(s.nrows):
                    values  = []
                    for col in range(s.ncols):
                        values.append(s.cell(row,col).value)
                        #       print ','.join(values)
                        #       print
                    rows.insert(    row,    values)
                self.sheets.insert( self.sheetIndex, rows)
            #
            # file load successful
            #
            #   clone output workbook
            #
            self.writeWorkBook  = copy(self.readWorkBook)
            #
            #   for garbage collection
            #
            self.readWorkBook   = None
            #
            # point at first sheet, row
            #
            self.sheetIndex = 0
            self.rowIndex   = 0
            self.excelInputFileName  = fName
            self.excelOutputFileName = self.excelInputFileName
            self.setTreeDirty(  False)
            self.setSheetsDirty( False)
            self.setRowDirty(   False)
            #
            self.addRecentFile( self.excelInputFileName)
            #
            #   guess output file names
            #
            #   excel output file name is same as excel input file name
            #   base name is input file name without .xls extention
            #   kml output file name is base name with .kml appended
            #
            if self.excelInputFileName.lower().endswith(".xls"):
                dotIndex    = self.excelInputFileName.lower().rfind(".xls")
                self.baseFileName   = self.excelInputFileName[0:dotIndex]
            else:
                self.baseFileName   = self.excelInputFileName
            #    
            self.kmlOutputFileName   = self.baseFileName + ".kml"    
            #
            #   label in gui
            #
            baseName = os.path.basename( self.excelInputFileName)
            dirName  = os.path.dirname(  self.excelInputFileName)
            #   gui label is two lines long, with line break
            self.fileNameLabel.setText( "File: %s\n%s" % ( dirName, baseName))
            #  
            self.updateSheetUi()
            self.updateRowUi()
            #
            self.updateExcelActions()
            self.updateKmlActions()
            #
            message = "Loaded %s" % baseName
            self.updateStatus( message)
            #
            if not  self.validateColumns():
                #
                # validateColumns will leave sheet and row 
                # indexes set at first invalid sheet and row, 
                # produce error message in status
                #
                self.updateSheetUi()
                self.updateRowUi()
            #
            return
        
    def excelWriteData( self):
        try:
            outputFile  = open( self.excelOutputFileName, mode="wb")
        
        except OSError, err:
            updateStatus( "open error %s on output file: %s" % (err.strerror, err.filename))
            return  False
        
        returnCode  = True    
        
        try:
            self.writeWorkBook.save( outputFile)
            
        except OSError, err:
            updateStatus( "write error %s on output file: %s" % (err.strerror, err.filename))
            returnCode = False
        
        try:
            outputFile.close()
            
        except OSError, err:
            updateStatus( "close error %s on output file: %s" % (err.strerror, err.filename))
            returnCode = False
        
        return  returnCode
            
##################################################################################################
    def helpAbout( self):
        QMessageBox.about( self, "About excel to kml",
                    """<b>excel to kml converter</b> v %s
                    <p>Copyright &copy; 2012 WorksOfEvil.com
                    All rights reserved. </p>
                    <p>This application can be used to convert some
                    Excel .xml files to Google Earth/Maps .kml files</p>
                    <p>Python %s - Qt %s - PyQt %s on %s""" % (
                    __version__, platform.python_version(),
                    QT_VERSION_STR, PYQT_VERSION_STR,
                    platform.system()))
        return

    def helpHelp( self):
        #hForm = helpform.HelpForm( ":/Help/index.html", self)
        hForm = helpform.HelpForm( "Help/index.html", self)
        hForm.show()
        return
        
    def isRowDirty( self):
        if  self.rowDirty:
            return True
        #
        return False

    def isSheetsDirty( self):
        if  self.SheetsDirty:
            return True
        #
        return False

    def isTreeDirty( self):
        if  self.treeDirty:
            return True
        #
        return False

###############################################################################
#
#   excel sheet to kml tree functions
#
#   see pykml and lxml library documentation
#
    def kmlDocFromSheets( self,
                          sheetNames,
                          sheetIconIndexes, 
                          sheets):
        #
        #   excel sheets to kml doc
        #   in general, build doc from inside out,
        #   bottom to top
        #
        #   style maps are an exception, top and first
        #   each style map is a three element list
        #
        #   string elements of style maps into a single list
        #   style maps must be a flat list
        #
        styleMaps      = []
        styleMapNames  = []
        sheetIconNames = []
        #
        for styleMapIndex in sheetIconIndexes:
            # 
            thisStyleMapName = self.kmlFolderIcons[styleMapIndex][0]
            sheetIconNames.append( thisStyleMapName)
            #
            # build a style map for each new style used       
            #
            if not thisStyleMapName in styleMapNames:
                styleMapNames.append( thisStyleMapName)
                thisStyleMapUrl = self.kmlFolderIcons[styleMapIndex][2]
                thisStyleMap    = self.kmlStyleMap( thisStyleMapUrl,
                                                    thisStyleMapName)
                #
                #   kmlStyleMap returns a 3 element list
                #   styleMaps must be a flat list
                #
                for sm in thisStyleMap:        
                    styleMaps.append( sm)
        #
        #   folders with placemarks, inside
        #            
        sheetFolders = self.kmlFolders( sheetNames, sheetIconNames, sheets)
        #
        #    for sf in sheetFolders:
        #        print etree.tostring( sf, pretty_print=True)
        #
        #   container folder, outside
        #
        #   popup dialog to get description text
        #
        (containerDescriptionText, ok) = QInputDialog.getText( self,
                                                               "Kml Container Folder",
                                            "Input a Description for\n the Container Folder:")
        #
        #containerFolder = self.kmlContainerFolder(  'Horse Camps',
        #                                            'Overnight Horse Camps in Kentucky')
        #
        baseName = os.path.basename( self.baseFileName)
        #
        containerFolder = self.kmlContainerFolder(  baseName,
                                                    containerDescriptionText)
        #
        #print "\nContainerFolder-before"
        #print etree.tostring( containerFolder, pretty_print=True)
        #
        #   put sheet folders into container
        #
        for sf in sheetFolders:
            containerFolder.append( sf)
        #print "\nContainerFolder-after"
        #print etree.tostring( containerFolder, pretty_print=True)
        #
        #   document header, outside
        #
        headerText  = '%s Kml File' % baseName
        #             
        header  = self.kmlDocHeader( headerText)
        #print "\nhead-before stylemaps"    
        #print etree.tostring( header, pretty_print=True)
        #
        #   append style maps to header
        #                            
        for sm in styleMaps:
            header.append( sm)                                  
        #print "\nhead-after stylemaps, before folders"    
        #print etree.tostring( header, pretty_print=True)
        #
        #   append container folder to header
        #    
        header.append( containerFolder)
        #
        #   insides are complete
        #    
        #print "\nhead-after folders"    
        #print etree.tostring( header, pretty_print=True)
        #
        #   doc wrapper, outside
        #                                  
        doc = self.kmlDocWrapper()
        #print "\nDocument-before"    
        #print etree.tostring( doc, pretty_print=True)
        #
        #   put header into doc
        #    
        doc.append( header)       
        #print "\nDocument-after"    
        #print etree.tostring( doc, pretty_print=True)
        #            
        return doc

    def kmlContainerFolder( self,
                            name, 
                            description):
        #
        #   return a container folder
        #   with name, description
        #    
        folder  = KML.Folder(
            KML.name( name),
            KML.open('1'),
            KML.description( description))
        #print etree.tostring(folder, pretty_print=True)    
        return  folder
                
    def kmlFolder(  self,
                    sheetName, 
                    sheetIconName, 
                    rows):
        #
        #   return a folder from a sheet,
        #   containing placemarks from
        #   the rows of the sheet
        #    
        placemarks  = self.kmlPlacemarks( sheetIconName, 
                                          sheetName,
                                          rows)
        #
        folder  = self.kmlContainerFolder( sheetName,
                                           sheetIconName)
        for pm in placemarks:
            folder.append( pm)
        #print etree.tostring(folder, pretty_print=True)    
        return  folder

    def kmlFolders( self,
                    sheetNames,  
                    sheetIconNames, 
                    sheets):
        #
        #   return a list of folders from a list of sheets
        #   each folder contains placemarks 
        #   from the rows of each sheet 
        #    
        folders = []
        #
        for sheetIndex in range( len(sheetNames)):
            rows          = sheets[sheetIndex]
            sheetName     = sheetNames[sheetIndex]
            sheetIconName = sheetIconNames[sheetIndex]
            #
            folder  = self.kmlFolder( sheetName, 
                                      sheetIconName, 
                                      rows)
            #
            folders.append( folder)      
            #
        return  folders

    def kmlDocWrapper( self):
        doc = KML.kml( )
        return  doc

    def kmlDocHeader(   self,
                        docName):
        docHeader = KML.Document(
                        KML.name( docName),
                        KML.open('1'))
        return  docHeader

    def kmlPlacemark(   self,
                        sheetIconName, 
                        row,
                        sheetName,
                        rowNumber):
        #
        #   return a single placemark 
        #   from a single excel row
        #   
        #   check row for valid lat and long data
        #
        if not self.validateRowGeoCode( row, sheetName, rowNumber):          
            return  None
        #
        latitudeString  = row[self.latitudeColumnIndex]
        longitudeString = row[self.longitudeColumnIndex]
        # 
        #   full description is the popup in the placemark
        #           
        #   it is the concatination of several cols in the excel data
        #
        fullDescription = []
        if not row[self.descriptionColumnIndex].isspace(): 
            fullDescription.append( row[self.descriptionColumnIndex])
        if not row[self.websiteColumnIndex].isspace():
            fullDescription.append( row[self.websiteColumnIndex])
        if not row[self.streetColumnIndex].isspace():
            fullDescription.append( row[self.streetColumnIndex])
        #
        #   format like a postal address
        #    
        townStateZip    = []    
        townStateZipString  = ""
        if not row[self.townColumnIndex].isspace():
            townStateZip.append( row[self.townColumnIndex])
        if not row[self.stateColumnIndex].isspace():
            townStateZip.append( row[self.stateColumnIndex])
        if not row[self.zipColumnIndex].isspace():
            townStateZip.append( row[self.zipColumnIndex])
        if 0 < len( townStateZip):
            townStateZipString  = " ".join( townStateZip)
            fullDescription.append( townStateZipString)
                
        if not row[self.telephoneColumnIndex].isspace():
            fullDescription.append( row[self.telephoneColumnIndex])
        fullDescriptionString   = "<br>\n".join( fullDescription)
        
        placemark = KML.Placemark(
            KML.name( row[self.nameColumnIndex]),
            KML.description( fullDescriptionString),
            KML.LookAt(
                KML.longitude( longitudeString),
                KML.latitude(  latitudeString),
                KML.altitude('0'),
                KML.heading('0'),
                KML.tilt('0'),
                KML.range('250.0'),
                KML.altitudeMode('relativeToGround'),
                GX.altitudeMode('relativeToSeaFloor'),
                ),
            KML.styleUrl('#'+ sheetIconName),
            KML.Point(
                KML.coordinates( longitudeString + ','
                               + latitudeString + ',0'),
                ),
            )
        
        return  placemark
        
    def kmlPlacemarks(  self,
                        sheetIconName, 
                        sheetName,
                        rows):
        #
        #   return a list of placemarks
        #   from a list of rows
        #
        placemarks  = []
        #
        #   skip first row with column names
        #    
        for rowIndex in range( 1, len(rows)):
            #
            thisRow   = rows[rowIndex]
            rowNumber = rowIndex + 1
            #
            placemark = self.kmlPlacemark( sheetIconName, 
                                           thisRow,
                                           sheetName,
                                           rowNumber)
            #   is placemark valid ?
            if( placemark is not None):
                placemarks.append( placemark)
            #        
        return  placemarks
                                
    def kmlStyleMap( self,
                     styleUrl, 
                     styleName):
        #
        #   produce a style map, and two syle icons,
        #   one for normal, one for selected.
        #   style map is a pair, with the two icons,
        #   given the id=styleName.
        #   each icon is the same image, styleUrl
        #   with the selected slightly larger than the 
        #   normal. the icons have the id's
        #   styleName0 and styleName1, used
        #   in the style map pair   
        #
        #   returned style map is a list with 3 elements
        #
        styleMap    = []
        styleMap.append( KML.StyleMap(
                            KML.Pair(
                                KML.key('normal'),
                                KML.styleUrl('#' + styleName + '0'),
                            ),
                            KML.Pair(
                                KML.key('highlight'),
                                KML.styleUrl('#' + styleName + '1'),
                            ),
                            id= styleName,
                            ))
        styleMap.append( KML.Style(
                            KML.IconStyle(
                                KML.scale('1.1'),
                                KML.Icon(
                                    KML.href( styleUrl),
                                ),
                                KML.hotSpot(  
                                    x="20",
                                    y="2",
                                    xunits="pixels",
                                    yunits="pixels",
                                ),
                            ),
                            id= styleName + "0",
                        ))
        styleMap.append( KML.Style(
                            KML.IconStyle(
                                KML.scale('1.3'),
                                KML.Icon(
                                    KML.href( styleUrl),
                                ),
                                KML.hotSpot(  
                                    x="20",
                                    y="2",
                                    xunits="pixels",
                                    yunits="pixels",
                                ),
                            ),
                            id= styleName + "1",
                        ))
        #print etree.tostring( styleMap[0],pretty_print=True)
        #print etree.tostring( styleMap[1],pretty_print=True)
        #print etree.tostring( styleMap[2],pretty_print=True)
        return styleMap                                
#
###############################################################################
#
#   kml file functions
#
    def kmlFileSave(   self):
        if( self.sheets is None
        or  self.kmlDoc is None):
            self.updateStatus(  "No Kml data to Save")
            return
            
        if self.kmlOutputFileName is None:
            self.kmlFileSaveAs()
            return
        #    
        if self.kmlWriteData():
            self.setTreeDirty( False)
            self.updateStatus(  "Kml data saved as %s" % self.kmlOutputFileName)
            return
        #
        self.updateStatus(  "Failed to save Kml data file %s" % self.kmlOutputFileName)
        #
        return

    def kmlFileSaveAs( self):
        if( self.sheets is None
        or  self.kmlDoc is None):
            self.updateStatus(  "No Data to Save")
            return
        dirName = ( os.path.dirname(self.kmlOutputFileName) 
                        if self.kmlOutputFileName is not None else ".")
        fName   = unicode(  QFileDialog.getSaveFileName( self,
                            "Google Earth/Maps Files - Save Output",    
                            dirName, "Earth files (*.kml *.KML)"))
        if not fName:
            return
        if "." not in fName:
            fName   += ".kml"
        #
        self.kmlOutputFileName   = fName 
        self.kmlFileSave()
        
        return
        
    def kmlWriteData( self):
        try:
            outputFile  = open( self.kmlOutputFileName, mode="wt")
        
        except OSError, err:
            updateStatus( "open error %s on output file: %s" % (err.strerror, err.filename))
            return  False
        
        returnCode  = True    
        
        try:
            outputFile.write( etree.tostring( self.kmlDoc, pretty_print=True))
            
        except OSError, err:
            updateStatus( "write error %s on output file: %s" % (err.strerror, err.filename))
            returnCode = False
        
        try:
            outputFile.close()
            
        except OSError, err:
            updateStatus( "close error %s on output file: %s" % (err.strerror, err.filename))
            returnCode = False
        
        return  returnCode
###########################################################################################        

    def loadInitialFile(    self):
        settings    = QSettings()
        fName   = unicode(  settings.value( "LastFile").toString())
        if fName and QFile.exists( fName):
            self.excelLoadFile( fName)

    def loadRecursion(  self, ancestorKey, ancestorWidgetItem, kmlChildren):
        #
        #   recursive function to fill tree widget
        #
        for kmlChild in kmlChildren:
            #
            #   traverse across width of tree
            #
            #   remove namespace prefix
            #
            kmlChildSplitTag    = kmlChild.tag.split( "}")
            kmlChildKey = QString(  kmlChildSplitTag[1])
            kmlChildId  = kmlChild.get( "id")
            if kmlChildId is not None:
                kmlChildKey.append( "\nid=")
                kmlChildKey.append( kmlChildId)
            widgetQStringList   = QStringList( kmlChildKey)
            
            if kmlChild.text is None:
                kmlChildText = None
            else:
                kmlChildText = QString( kmlChild.text)
                                                
            if (ancestorKey.startsWith( "Placemark")
            and kmlChildKey.startsWith( "description")):
                #
                #   Placemark description items are embedded html,
                #   quoted in a cdata section. They appear in a 
                #   popup on the map web page or googleEarth
                #
                childWidgetItem = QTreeWidgetItem(  ancestorWidgetItem,
                                                    widgetQStringList)
                descriptionWidget   = QWebView()
                self.treeWidget.setItemWidget( childWidgetItem,
                                               1,
                                               descriptionWidget)
                descriptionWidget.setAutoFillBackground( True)
                descriptionWidget.setHtml( kmlChildText)    
                descriptionWidget.setMaximumHeight( 175)
            else:
                #
                #   other items are simple
                #
                if kmlChildText is not None:
                    widgetQStringList.append( kmlChildText)
                childWidgetItem = QTreeWidgetItem(  ancestorWidgetItem,
                                                    widgetQStringList)
                childWidgetItem.setTextAlignment(   1,
                                            (Qt.AlignJustify | Qt.AlignTop))
            #
            #   recurse down tree branch
            #
            kmlGrandChildren = kmlChild.getchildren()
            #
            if 0 < len(kmlGrandChildren):
                self.loadRecursion( kmlChildKey, childWidgetItem, kmlGrandChildren)
        #
        return

    def lookupGeoCode(  self):
        #
        #   call Google web service to get 
        #   latitude, longitude from postal address
        #
        postalAddress = self.streetEdit.text() + ", "\
                      + self.townEdit.text()   + ", "\
                      + self.stateEdit.text()  + " "\
                      + self.zipEdit.text()
        #
        try:
            place, (lat, lng) = self.googleGeoCoder.geocode( postalAddress)
        except geocoders.google.GQueryError, exData:
            self.updateStatus(  "GeoCoding Data Error: %s" % exData)
            return False
        except geocoders.google.GTooManyQueriesError, exData:
            self.updateStatus(  "GeoCoding Too Many Queries Today: %s" % exData)
            return False
        except geocoders.base.GeocoderResultError, exData:
            self.updateStatus(  "GeoCoding Error: %s" % exData)
            return False
        except  ValueError, exData:
            self.updateStatus(  "GeoCoding Error: %s" % exData)
            return False           
        #
        latitudeString  = "% .6f" % lat
        longitudeString = "% .6f" % lng
        #
        self.updateStatus(  "GeoCoded: %s" % place)
        #
        #print   "%s: %.5 %.5" % (place, lat, lng)
        #
        self.latitudeEdit.setText(  latitudeString)
        self.longitudeEdit.setText( longitudeString)
        #
        self.setRowDirty( True)
        #
        return  True
                        
    def nextRow(   self):
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex < 0):
            return
        topRowIndex    = len( self.sheets[self.sheetIndex]) - 1
        if self.rowIndex >= topRowIndex: 
            return
            
        self.rowIndex   += 1
        #
        self.updateRowUi()
        return
        
    def nextSheet(   self):
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex < 0):
            return
            
        topSheetIndex    = len( self.sheets) - 1
        if self.sheetIndex >= topSheetIndex: 
            return
            
        self.sheetIndex += 1
        #   first row in new sheet
        self.rowIndex   = 0
        #
        self.updateSheetUi()
        self.updateRowUi()
        return

    def okToContinue(   self):
        #
        #   maybe clean, maybe one or other dirty, maybe both
        #
        #   False return means at least one file was dirty
        #   and reply to MessageBox was "Cancel"
        #
        if self.isTreeDirty():
            kmlReply   = QMessageBox.question( self,
                        "Excel to Kml Unsaved Changes",
                        "Save unsaved kml file changes",
                        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if kmlReply == QMessageBox.Cancel:
                return  False
            if kmlReply == QMessageBox.Yes:
                self.kmlFileSave()
            # kmlReply was No or file was saved    

        if self.isSheetsDirty():
            excelReply   = QMessageBox.question( self,
                        "Excel to Kml Unsaved Changes",
                        "Save unsaved Excel file changes",
                        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if excelReply == QMessageBox.Cancel:
                return  False
            if excelReply == QMessageBox.Yes:
                self.excelFileSave()
            # excelReply was No or file was saved    
        #
        #   True return means files were clean,
        #   or some combination of "No" replies 
        #   or files saved 
        #    
        return True

    def populateTreeWidget( self,   kmlTree):
        #
        self.treeWidget.clear()
        self.treeWidget.setColumnCount( 2)
        self.treeWidget.setHeaderLabels( [  "Keys", "Text"])
        self.treeWidget.setWordWrap( True)
        #
        if ( kmlTree is None):
            return
        #
        selected = None                
        #
        #   remove namespace prefix
        #
        splitTag = kmlTree.tag.split( "}")
        ancestorKey = QString(  splitTag[1])
        ancestorId  = kmlTree.get( "id")
        if ancestorId is not None:
            ancestorKey.append( "\nid=")
            ancestorKey.append( ancestorId)
        ancestorWidgetQStringList   = QStringList( ancestorKey)
        #
        if kmlTree.text is None:
            ancestorText    = None
        else:
            ancestorText    = kmlTree.text
            ancestorWidgetQStringList.append( ancestorText)
        #    
        ancestorWidgetItem  = QTreeWidgetItem(  self.treeWidget,
                                                ancestorWidgetQStringList)
        ancestorWidgetItem.setTextAlignment(    1,
                                            (Qt.AlignJustify | Qt.AlignTop))
        kmlChildren = kmlTree.getchildren()
        #
        if 0 < len(kmlChildren):
            self.loadRecursion(  ancestorKey, ancestorWidgetItem, kmlChildren)
        #
        self.resizeTreeWidgetCol()
        #
        return
            
    def prevRow(   self):
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex <= 0):
            return
        #
        self.rowIndex   -= 1
        #
        self.updateRowUi()
        #
        return
        
    def prevSheet(   self):
        if( self.sheets is None
        or  self.sheetIndex <= 0):
            return
        #    
        self.sheetIndex  -= 1
        #   first row on new sheet
        self.rowIndex    = 0
        #
        self.updateSheetUi()
        self.updateRowUi()
        #
        return
        
    def resizeTreeWidgetCol( self):
        #   resize the key column horizontally
        self.treeWidget.resizeColumnToContents( 0)
        #
        return

    def setRowDirty( self, setTo):
        #
        #   check state for change
        #
        if setTo:
            if  self.isRowDirty():
                # no change
                return True
        else:
            if not self.isRowDirty():
                # no change
                return False
        #
        #   row dirty has changed
        #
        if setTo:
            self.rowDirty   = True
            #
            #   disable file actions, data actions
            #   and sheet and row change pushbuttons
            #   enable row undo and update pushbuttons
            self.updateRowAndSheetPushbuttons()
            #
            return True
        #
        self.rowDirty   = False
        #
        #   enable file actions, data actions
        #   and sheet and row change pushbuttons
        #   enable row undo and update pushbuttons
        #
        #   update ui's will check row dirty
        #   and enable-disable pushbuttons
        #
        #   bad idea, modified change signal on description
        #   will cause recursion
        #   put pushbutton logic another function
        #
        self.updateRowAndSheetPushbuttons()
        #
        return  False

    def setSheetsDirty( self, setTo):
        #
        #   check state for change
        #
        if setTo:
            if  self.isSheetsDirty():
                # no change
                return True
        else:
            if not self.isSheetsDirty():
                # no change
                return False
        #
        #   sheet dirty has changed
        #
        if setTo:
            self.SheetsDirty   = True
            #
            #   enable excel file save action
            #   excel fileSaveAs action is enabled whenever
            #   sheet data, output workbook are not none
            #
            self.updateExcelActions()
            #
            return True
            #
        self.SheetsDirty   = False
        #
        #   disable excel file save action
        #
        self.updateExcelActions()
        #
        return False

    def setTreeDirty( self, setTo):
        #
        #   check state for change
        #
        if setTo:
            if  self.isTreeDirty():
                # no change
                return True
        else:
            if not self.isTreeDirty():
                # no change
                return False
        #
        #   tree dirty has changed
        #
        if setTo:
            self.treeDirty   = True
            #
            #   enable kml file save action
            #   kml fileSaveAs action is enabled whenever tree is not none
            self.updateKmlActions()
            #
            return True
        #
        self.treeDirty   = False
        #
        #   disable kml file save action
        #
        self.updateKmlActions()
        #
        return False
        
    def updateExcelActions( self):
        #
        #   check for data
        #
        if self.sheets is None:
            self.excelFileSaveAction.setEnabled(   False)
            self.excelFileSaveAsAction.setEnabled( False)
            return
        #    
        self.excelFileSaveAsAction.setEnabled( True)
        #    
        if self.isSheetsDirty():
            self.excelFileSaveAction.setEnabled(   True)
        else:
            self.excelFileSaveAction.setEnabled(   False)
        #    
        return
        
    def updateExcelMenu( self):
        #
        #   this is tricked up to display file lru list
        #
        self.excelMenu.clear()
        self.addActions( self.excelMenu,
                         self.excelMenuActions)
                         #self.excelMenuActions[:-1])
        current = QString( self.excelInputFileName) \
                    if self.excelInputFileName is not None else None
        recentFiles = []
        for fName in self.recentFiles:
            if fName != current and QFile.exists( fName):
                recentFiles.append( fName)
        if recentFiles:
            self.excelMenu.addSeparator()
            for i, fName in enumerate( recentFiles):
                action  = QAction(  QIcon( ":/Icon/excel.png"),
                                    "&%d %s" % (
                                    i + 1, QFileInfo( fName).fileName()), self)
                action.setData( QVariant( fName))
                self.connect(   action, SIGNAL( "triggered()"),
                                self.excelLoadFile)
                self.excelMenu.addAction( action)
    
    def updateKmlActions( self):
        #
        #   check for kml data
        #
        if self.kmlDoc is None:
            self.kmlFileSaveAsAction.setEnabled( False)
            self.kmlFileSaveAction.setEnabled(   False)
        else:
            self.kmlFileSaveAsAction.setEnabled( True)
            if self.isTreeDirty():
                self.kmlFileSaveAction.setEnabled( True)
            else:
                self.kmlFileSaveAction.setEnabled( False)
        #    
        #   check for Excel data
        #
        if self.sheets is None: 
            self.dataKmlAction.setEnabled( False)
            self.dataGeoAction.setEnabled( False)
        else:
            self.dataKmlAction.setEnabled( True)
            self.dataGeoAction.setEnabled( True)
        #
        return
    
    def undoRowData( self):
        #
        #   replace data in edits with data from 3d array 
        #
        self.setRowDirty(   False)
        self.updateStatus( "Row %d changes undone" % (self.rowIndex + 1))
        #
        self.updateRowUi()
        #
        return
        
    def updateRowAndSheetPushbuttons(   self):
        #
        #   this is factored out of updateRowUi
        #   because calling updateRowUi from
        #   updateRowData was causing a mutual 
        #   recursion by toggling the modified
        #   indicator on the description edit   
        #
        #   calling this independantly should
        #   prevent the recursion
        #
        #   see notes in updateRowData and
        #   updateRowUi
        #
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex < 0):
            #   no sheet data, disable all buttons
            self.lookupGeoCodeButton.setEnabled( False) 
            self.prevRowButton.setEnabled(       False)
            self.undoRowButton.setEnabled(       False)
            self.updateRowButton.setEnabled(     False)
            self.nextRowButton.setEnabled(       False)
            self.prevSheetButton.setEnabled(     False)
            self.nextSheetButton.setEnabled(     False)
            self.sheetIconDropdown.setEnabled(   False)
            #
            return
        #
        #   this enables if we are on a data row, 
        #   the lookup will convert street, city, state, zip
        #   address to longitude and latitude, if address
        #   is ccomplete and non-ambigious
        #
        if 0 < self.rowIndex:
            self.lookupGeoCodeButton.setEnabled( True)
        else:
            self.lookupGeoCodeButton.setEnabled( False)
        #    
        if self.isRowDirty():
            #
            #   these enable when rowDirty is true
            #   after a line edit is changed
            self.undoRowButton.setEnabled(   True)
            self.updateRowButton.setEnabled( True)
            #   prevent changing rows or sheets
            #   until changes are updated or undone
            self.prevRowButton.setEnabled(   False)
            self.nextRowButton.setEnabled(   False)            
            #
            self.prevSheetButton.setEnabled(   False)
            self.nextSheetButton.setEnabled(   False)
            self.sheetIconDropdown.setEnabled( False)        
            #
            return
        #     
        #   no row changes to update
        #
        self.undoRowButton.setEnabled(   False)
        self.updateRowButton.setEnabled( False)
        #
        #   is there a previous row ?
        #
        if 0 < self.rowIndex:
            self.prevRowButton.setEnabled( True)
        else:
            self.prevRowButton.setEnabled( False)
        #
        #   is there a next row ?
        #
        if (len( self.sheets[self.sheetIndex]) - 1) > self.rowIndex:
            self.nextRowButton.setEnabled( True)
        else:
            self.nextRowButton.setEnabled( False)
        #
        #   is there a previous sheet ?
        #    
        if 0 < self.sheetIndex:
            self.prevSheetButton.setEnabled( True)
        else:
            self.prevSheetButton.setEnabled( False)
        #
        #   is there a next sheet ?
        #    
        if (len( self.sheets) - 1) > self.sheetIndex:
            self.nextSheetButton.setEnabled( True)
        else:
            self.nextSheetButton.setEnabled( False)
        #
        self.sheetIconDropdown.setEnabled( True)
        #
        return
        
    def updateRowData( self):
        #
        #   we know at least one edit control contains different 
        #   data than the sheet array, but we do not know which one
        #   check them all
        #
        #   see notes in updateRowAndSheetPushbuttons 
        #   and updateRowUi about recursion
        #
        #   update row data does not need to call updateRowUi after changing 
        #   the sheet data because the changes are already in the visible edits. 
        #   in fact, that is where the updated data comes from
        #
        thisSheetDirty = False
        #
        #   write workbook was cloned from read workbook when file was read
        #
        sheetIndex   = self.sheetIndex
        thisWwbSheet = self.writeWorkBook.get_sheet( sheetIndex)
        #
        #   this row is in 3d data array
        #
        rowIndex = self.rowIndex
        thisRow  = self.sheets[ sheetIndex][ rowIndex]
        #
        newData = unicode( self.nameEdit.text())
        columnIndex = self.nameColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #    
        newData = unicode( self.descriptionTextEdit.toPlainText())
        columnIndex = self.descriptionColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.streetEdit.text())
        columnIndex = self.streetColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.townEdit.text())
        columnIndex = self.townColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.stateEdit.text())
        columnIndex = self.stateColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.zipEdit.text())
        columnIndex = self.zipColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.telephoneEdit.text())
        columnIndex = self.telephoneColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.websiteEdit.text())
        columnIndex = self.websiteColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.latitudeEdit.text())
        columnIndex = self.latitudeColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        newData = unicode( self.longitudeEdit.text())
        columnIndex = self.longitudeColumnIndex
        if( newData != thisRow[columnIndex]): 
            thisRow[columnIndex] = newData
            thisWwbSheet.write( rowIndex, columnIndex, newData)
            thisSheetDirty = True
        #
        #   sheetsDirty means data should be written out
        #    
        if thisSheetDirty:
            self.setSheetsDirty( True)
        #
        self.setRowDirty(   False)
        self.updateStatus( "Row %d updated" % (rowIndex + 1))
        #
        return
        
    def updateRowUi(  self):
        #
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex < 0):
            self.nameEdit.setText(  " ")
            self.descriptionTextEdit.setPlainText(  " ")
            self.descriptionTextEdit.document().setModified(   False) 
            self.streetEdit.setText(    " ")
            self.townEdit.setText(  " ")
            self.stateEdit.setText(  " ")
            self.zipEdit.setText(  " ")
            self.telephoneEdit.setText( " ")
            self.websiteEdit.setText(   " ")
            self.latitudeEdit.setText(  " ")
            self.longitudeEdit.setText( " ")
            self.rowNofMlabel.setText( "Row 0 of 0")
            #
            self.setRowDirty( False)
            #
            self.updateRowAndSheetPushbuttons()
            #
            return
        #    
        self.rowNofMlabel.setText( "Row %d of %d" % 
            ((self.rowIndex + 1), len( self.sheets[self.sheetIndex])))
        #
        if self.isRowDirty():
            #
            #   if rowDirty is true, at least one
            #   of the edit controls contains
            #   different text than the sheet array
            #   do not update the edit controls,
            #   as this would overwrite the new data
            #
            self.updateRowAndSheetPushbuttons()
            #
            return
        #
        row = self.sheets[self.sheetIndex][self.rowIndex]
        #
        #   update edit controls from sheet array
        #
        self.nameEdit.setText(      row[self.nameColumnIndex])
        self.descriptionTextEdit.setPlainText( row[self.descriptionColumnIndex])
        self.descriptionTextEdit.document().setModified(   False) 
        self.streetEdit.setText(    row[self.streetColumnIndex])
        self.townEdit.setText(      row[self.townColumnIndex])
        self.stateEdit.setText(     row[self.stateColumnIndex])
        self.zipEdit.setText(       row[self.zipColumnIndex])
        self.telephoneEdit.setText( row[self.telephoneColumnIndex])
        self.websiteEdit.setText(   row[self.websiteColumnIndex])
        self.latitudeEdit.setText(  row[self.latitudeColumnIndex])
        self.longitudeEdit.setText( row[self.longitudeColumnIndex])
        #
        self.setRowDirty( False)
        #
        self.updateRowAndSheetPushbuttons()
        #
        return            

    def updateSheetIconIndex( self, newIndex):
        if( self.sheets is None
        or  newIndex < 0
        or  newIndex >= 8):
            #   no data or invalid index
            return
        #
        self.sheetIconIndexes[self.sheetIndex] = newIndex
        #
        return
        
    def updateSheetUi(  self):
        if( self.sheets is None
        or  self.sheetIndex < 0
        or  self.rowIndex < 0):
            #
            self.sheetNameLabel.setText( "Sheet: None")
            self.sheetNofMlabel.setText( "Sheet 0 of 0")
            self.sheetIconDropdown.setCurrentIndex(0)
            #
        else:
            #    
            self.sheetNameLabel.setText( "Sheet: %s" % self.sheetNames[self.sheetIndex])
            self.sheetNofMlabel.setText( "Sheet %d of %d" % 
                                            ((self.sheetIndex+1), len( self.sheetNames)))
            #
            self.sheetIconDropdown.setCurrentIndex( self.sheetIconIndexes[self.sheetIndex])
        #
        self.updateRowAndSheetPushbuttons()
        #
        return
                                        
    def updateStatus(   self, message):
        #
        self.statusBar().showMessage(   message, 5000)
        #
        self.logListWidget.addItem(    message)
        self.logListWidget.setCurrentRow( self.logListWidget.count() - 1)
        #
        self.setWindowModified( self.isTreeDirty()
                             or self.isSheetsDirty()
                             or self.isRowDirty())

        if self.excelInputFileName is not None:
            self.setWindowTitle(    "excel to kml - %s[*]" % \
                                    os.path.basename( self.excelInputFileName))
            return

        self.setWindowTitle(        "excel to kml [*]")
        return

    def validateColumns( self):
        if( self.sheets is None):
            #   no data
            self.updateStatus( "no data to validate columns") 
            return True
        #
        #   first row is column headings
        #   an error will leave sheet index
        #   pointed at sheet in error
        #   produce an error message in status
        #
        self.colsValid  = True
        #
        for self.sheetIndex in range( len( self.sheets)):
            
            headerRow = self.sheets[self.sheetIndex][0]
        
            if ( headerRow[self.nameColumnIndex].lower()         
            != "name"):
                self.colsValid  = False
                colErrorString  = "Name"
                break
            if ( headerRow[self.descriptionColumnIndex].lower()  
            != "description"):
                self.colsValid  = False
                colErrorString  = "Description"
                break
            if ( headerRow[self.streetColumnIndex].lower()       
            != "street"):
                self.colsValid  = False
                colErrorString  = "Street"
                break
            if ( headerRow[self.townColumnIndex].lower()         
            != "town"):
                self.colsValid  = False
                colErrorString  = "Town"
                break
            if ( headerRow[self.stateColumnIndex].lower()        
            != "state"):
                self.colsValid  = False
                colErrorString  = "State"
                break
            if ( headerRow[self.zipColumnIndex].lower()          
            != "zip"):
                self.colsValid  = False
                colErrorString  = "Zip"
                break
            if ( headerRow[self.telephoneColumnIndex].lower()    
            != "telephone"):
                self.colsValid  = False
                colErrorString  = "Telephone"
                break
            if ( headerRow[self.websiteColumnIndex].lower()      
            != "website"):
                self.colsValid  = False
                colErrorString  = "Website"
                break
            if ( headerRow[self.latitudeColumnIndex].lower()     
            != "latitude"):
                self.colsValid  = False
                colErrorString  = "Latitude"
                break
            if ( headerRow[self.longitudeColumnIndex].lower()    
            != "longitude"):
                self.colsValid  = False
                colErrorString  = "Longitude"
                break
        #        
        if self.colsValid: 
            #   reset sheet index to zero
            #   row index was not changed       
            self.sheetIndex = 0
            return  True      
        #
        #   columns are not valid
        #   error message should say
        #   which sheet, col in error
        #
        #
        if 0 > self.sheetIndex or len(self.sheetNames) <= self.sheetIndex:
            sheetErrorString = "Unknown"
        else:
            sheetErrorString = self.sheetNames[ self.sheetIndex]
        #            
        self.updateStatus( "%s column on sheet %s has incorrect label" 
                           % ( colErrorString, sheetErrorString))
        return False 

    def validateGeoCoding( self):
        #
        # the logic of this function is crap
        #
        # it tries to start with first data row,
        # scan the all rows on all sheets, say OK or not
        #
        # if a row has bad data, the scan will stop, put an
        # error message in status, and return False, so the 
        # caller will display the row in error
        #
        # the crap part is trying to stop at the first bad record,
        # to allow manual correction of that record, then expect 
        # the user to restart scanning from the that row, to the 
        # next error, or the end.
        #
        # if a row has an unfixable error, the user will have to
        # manually skip forward to the next record, and restart
        # the scan. The KmlPlacemark function also performs a check
        # and will not produce a placemark for an invalid geocode
        #
        if( self.sheets is None):
            #   no data
            self.updateStatus( "no data to check for GeoCoding") 
            return
        #
        #   guess we want to start checking at current row
        #
        startRowIndex   = self.rowIndex
        if 0 >= startRowIndex:
            startRowIndex = 1
        startSheetIndex = self.sheetIndex
        #    
        if( 0 >= startSheetIndex
        and 1 >= startRowIndex):
            #
            # we are starting at first data row on first sheet
            # self.geocodingValid is only true when all rows
            # have been checked
            #
            self.geocodingValid = True
        else:
            #
            #   we are not positioned at first sheet, first data row
            #   popup dialog box, ask user where to start
            #
            startRecordQuestion = QMessageBox()
            #
            startRecordQuestion.setIcon( QMessageBox.Question)
            startRecordQuestion.setText( "Where to Start GeoCode Check ?")
            startRecordQuestion.setInformativeText( 
              "At the First data row on the first sheet,\nor the Current row (row %d on sheet %d)"
                % ((startRowIndex + 1), (startSheetIndex + 1)))
            firstButton   = startRecordQuestion.addButton( 
                "&First",   QMessageBox.YesRole)
            currentButton = startRecordQuestion.addButton( 
                "&Current", QMessageBox.NoRole)
            cancelButton = startRecordQuestion.addButton( 
                "Cancel",  QMessageBox.RejectRole)
            startRecordQuestion.setDefaultButton( currentButton)
            # 
            startRecordQuestion.exec_()
            #
            if ( startRecordQuestion.clickedButton() == firstButton):
                # first row on first sheet
                startSheetIndex = 0
                startRowIndex   = 1
                self.geocodingValid = True
            elif(startRecordQuestion.clickedButton() == currentButton):
                # current row
                pass
            else:
                # cancel button
                return
        #    
        #   check all rows on all sheets for 
        #   numbers in lat, long cols
        #   an error will leave sheet and row indexes
        #   pointed at first sheet, row in error
        #   produce an error message in status
        #
        #   NOT the self value, just this function
        #
        geocodingValid  = True
        #
        for sheetIndex in range( startSheetIndex, len( self.sheets)):
            thisSheet = self.sheets[ sheetIndex]
            #   for error message in validateRowGeoCode
            sheetNameString = self.sheetNames[ sheetIndex]
            for rowIndex in range( startRowIndex, len( thisSheet)):   
                thisRow = thisSheet[ rowIndex]
                #   for error message in validateRowGeoCode
                rowNumber  = rowIndex + 1
                #
                if not self.validateRowGeoCode( thisRow, sheetNameString, rowNumber):
                    #
                    geocodingValid = False
                    errorSheetIndex = sheetIndex
                    errorRowIndex   = rowIndex
                    #
                    break
            if not geocodingValid:
                break
            #
            # after first pass thru inner loop
            # start next pass of inner loop
            # at first data record of next sheet
            #
            startRowIndex = 1
            #
                                
        if( geocodingValid
        and self.geocodingValid):
            #
            self.updateStatus( "all rows on all sheets have valid Latitude and Longitude") 
            #
            return True
        #
        #   set self.indexes to first sheet, first row in error
        #   validateRow set error message in status
        #
        if not geocodingValid:
            self.geocodingValid = False
            #   set position, ? caller will update row, sheet displays ?    
            self.sheetIndex = errorSheetIndex
            self.rowIndex   = errorRowIndex
            #    
            return False
            #
        #
        self.updateStatus( 
            "the Current row, and all rows after, have valid Latitude and Longitude") 
        #             
        return  True
                 
    def validateRowGeoCode( self, row, sheetName, rowNumber):
        #
        #   name and number are used in error messages
        #            
        #   this is NOT the global
        #
        geocodingValid = True
        #    
        longitudeString = row[self.longitudeColumnIndex] 
        #
        if((0 >= len( longitudeString))
        or longitudeString.isspace()):
            geocodingValid = False
            geoErrorString = "no Longitude"
        else:
            try:
                longitudeNumber = float( longitudeString)
                if(( 180.0 < longitudeNumber) 
                or (-180.0 > longitudeNumber)):
                    geocodingValid = False
                    geoErrorString = "invalid Longitude"
            except ValueError:
                geocodingValid = False
                geoErrorString = "invalid Longitude"
                
        if geocodingValid:
            latitudeString = row[self.latitudeColumnIndex] 
            if ((0 >= len( latitudeString))
            or latitudeString.isspace()):
                geocodingValid = False
                geoErrorString  = "no Latitude"
            else:
                try:
                    latitudeNumber = float( latitudeString)
                    if(( 90.0 < latitudeNumber) 
                    or (-90.0 > latitudeNumber)):
                        geocodingValid = False
                        geoErrorString = "invalid Latitude"
                except ValueError:
                    geocodingValid = False
                    geoErrorString = "invalid Latitude"
                    
        if geocodingValid:
            return True           
        #
        #   error message
        #
        self.updateStatus( "row %d on sheet %s has %s data" 
                           % ( rowNumber, sheetName, geoErrorString))
        return False
        
    def wireUpEvents(   self): 
        #
        #   hook aboutToShow for file menu, file menu is built dynamically
        #
        self.connect(   self.excelMenu,  SIGNAL( "aboutToShow()"),
                        self.updateExcelMenu)
        #
        #   pushbuttons
        #
        self.connect(   self.lookupGeoCodeButton, SIGNAL( "clicked()"),
                        self.lookupGeoCode)
        self.connect(   self.nextRowButton,       SIGNAL( "clicked()"),
                        self.nextRow)   
        self.connect(   self.nextSheetButton,     SIGNAL( "clicked()"),
                        self.nextSheet)
        self.connect(   self.prevRowButton,       SIGNAL( "clicked()"),
                        self.prevRow)                              
        self.connect(   self.prevSheetButton,     SIGNAL( "clicked()"),
                        self.prevSheet)           
        self.connect(   self.undoRowButton,       SIGNAL( "clicked()"),
                        self.undoRowData)           
        self.connect(   self.updateRowButton,     SIGNAL( "clicked()"),
                        self.updateRowData)           
        #
        #   tree widget
        #
        self.connect(   self.treeWidget,    SIGNAL( "expanded( QModelIndex)"),
                        self.resizeTreeWidgetCol)
        self.connect(   self.treeWidget,    SIGNAL( "collapsed( QModelIndex)"),
                        self.resizeTreeWidgetCol)
        #
        #   line edits
        #               
        self.connect(   self.nameEdit,            SIGNAL( "textEdited( QString)"),
                        self.changedNameData)
        self.connect(   self.descriptionTextEdit, SIGNAL( "modificationChanged( bool)"),
                        self.changedDescriptionData)
        self.connect(   self.streetEdit,          SIGNAL( "textEdited( QString)"),
                        self.changedStreetData)
        self.connect(   self.townEdit,            SIGNAL( "textEdited( QString)"),
                        self.changedTownData)
        self.connect(   self.stateEdit,           SIGNAL( "textEdited( QString)"),
                        self.changedStateData)
        self.connect(   self.zipEdit,             SIGNAL( "textEdited( QString)"),
                        self.changedZipData)
        self.connect(   self.telephoneEdit,       SIGNAL( "textEdited( QString)"),
                        self.changedTelephoneData)
        self.connect(   self.websiteEdit,         SIGNAL( "textEdited( QString)"),
                        self.changedWebsiteData)
        self.connect(   self.latitudeEdit,        SIGNAL( "textEdited( QString)"),
                        self.changedLatitudeData)
        self.connect(   self.longitudeEdit,       SIGNAL( "textEdited( QString)"),
                        self.changedLongitudeData)
        #
        #   combo box
        #
        self.connect(   self.sheetIconDropdown,   SIGNAL( "currentIndexChanged( int)"),
                        self.updateSheetIconIndex)
        #
        return

###########################################################################################        

def main():
    app = QApplication( sys.argv)
    app.setOrganizationName(    "WorksOfEvil")
    app.setOrganizationDomain(  "WorksOfEvil.com")
    app.setApplicationName(     "excel to kml")
    app.setWindowIcon(  QIcon(  ":/kml.png"))
    form    = MainWindow()
    form.show()
    app.exec_()

main()

