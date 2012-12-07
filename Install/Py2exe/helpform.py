#!/usr/bin/env python
#
#
from	__future__ import division
from	__future__ import print_function
from	__future__ import unicode_literals

from	future_builtins import *

from	PyQt4.QtCore	import (QUrl, Qt, SIGNAL, SLOT)
from	PyQt4.QtGui		import (QAction, QApplication, QDialog, QIcon,
								QKeySequence, QLabel, QTextBrowser,
								QToolBar, QVBoxLayout)
import  excelToKmlQrc

class	HelpForm( QDialog):

	def	__init__( self, page, parent=None):
		super( HelpForm, self).__init__( parent)
		self.setAttribute( Qt.WA_DeleteOnClose)
		self.setAttribute( Qt.WA_GroupLeader)

		backAction = QAction( QIcon( ":/Icon/back.png"), "&Back", self)
		backAction.setShortcut( QKeySequence.Back)
		homeAction = QAction( QIcon( ":/Icon/home.png"), "&Home", self)
		homeAction.setShortcut( "Home")
		self.pageLabel = QLabel()

		toolBar	= QToolBar()
		toolBar.addAction( backAction)
		toolBar.addAction( homeAction)
		toolBar.addWidget( self.pageLabel)
		self.textBrowser = QTextBrowser()
		
		layout	= QVBoxLayout()
		layout.addWidget( toolBar)
		layout.addWidget( self.textBrowser, 1)
		self.setLayout(	layout)

		self.connect(	backAction, SIGNAL( "triggered()"),
						self.textBrowser, SLOT( "backward()"))
		self.connect(	homeAction, SIGNAL( "triggered()"),
						self.textBrowser, SLOT( "home()"))
		self.connect(	self.textBrowser, SIGNAL( "sourceChanged(QUrl)"),
						self.updatePageTitle)

		self.textBrowser.setSearchPaths( [":/Help"])
		qrcEscapedPage  = "qrc:/" + page
		self.textBrowser.setSource( QUrl( qrcEscapedPage))
		self.resize( 400, 600)
		self.setWindowTitle( "{0} Help".format(
				QApplication.applicationName()))

	def	updatePageTitle( self):
		self.pageLabel.setText(	self.textBrowser.documentTitle())

if __name__ == "__main__":
	import	sys
	
	app		= QApplication(	sys.argv)
	form	= HelpForm( "index.html")
	form.show()
	app.exec_()

