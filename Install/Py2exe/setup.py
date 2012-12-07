#
#   run: python setup.py py2exe
#
from py2exe.build_exe import py2exe
from distutils.core import setup
setup(  windows = [ 'excelToKml.py'],
        zipfile = None,
        options={ "py2exe": { 
                    "includes": [ 'sip', 
                                  'PyQt4.QtNetwork',
                                  'lxml._elementpath'],
                     "bundle_files": 1
                            }
                }
    )