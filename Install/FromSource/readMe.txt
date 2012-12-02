File: readMe.txt of the ExcelToKml Install/FromSource directory

Installation instructions for Source form of the program. 
Simplist form of install, but requires technical knowledge.
Works on all platforms, Windows, Mac, Linux.

1) Install a version of Python 2.7 on your system. Free versions are availible for all platforms.
	
	for windows, use ActiveState python, community edition, MSI installer, 32 or 64 bit, at:
	http://www.activestate.com/activepython/downloads
	
	for mac, use ActiveState python, community edition, DMG image, at:
	http://www.activestate.com/activepython/downloads
	
	for linux, use your package manager, either apt-synaptic or rpm-yum.
	
	
2) Install required python imports. Python includes its own package installation system,
jokingly called "The Cheese Shop". For the active state version, it is called "Python Package Manager" or PyPM.
For other versions, it may be called easy_install, or setup. It is invoked by opening a command prompt,
or shell window and typing "pypm [package name]. 

If you are afraid to use the command prompt, you are a wimp, and should use another installation method.

On linux, use the system pacakage manager where possible, python package install only when the package is not
available from the system repositories. 

The required imports are:
	PyQt4	- Big package, GUI and event framework
	geopy	- Address to Longitude/Latitude lookup
	lxml	- basic xml file parse and generate
	xlrd	- read Excel files
	xlutils	- Excel file utilities
	pykml	- kml file generation
	
If the package manager reports the package is not found, check spelling and retry. All are present.

3) Copy the three python files in this directory to a convienient place on your system. 
I would recommend a new subdirectory of your home directory, so you have read/write permission. 
You will need write permission to save data files. 
The three files are:
	excelToKml.py		- Program file, class file for application
	excelToKmlQrc.py	- Resource file, icons and help pages
	helpform.py			- Program file, Class file for help form

4) Open a command prompt or shell window and start the program file, typing:
	excelToKml.py or ./excelToKml.py.
The purpose of the shell prompt is to see any error messages. These may be about missing imports, 
or other easily fixed problems.

5) Once the program can be started without errors, put a data file, with the correct column set, on the system.
After a data file has been read by the program, it will remember the last data file opened, and attempt to re-open
it, when the program is re-started.

Use the Operating System to produce a startup shortcut or icon for the program.
