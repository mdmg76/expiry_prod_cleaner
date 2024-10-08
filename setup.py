"""
This is a setup.py script generated by py2applet

Usage:
    python setup.py py2app
"""

from setuptools import setup

APP = ['gui_expiry_prod_cleaner.py']
DATA_FILES = ['aid.jpg','database.xlsx']
OPTIONS = {
    'argv_emulation': True,
    'iconfile':'aidIcon1.icns',
    'plist': {'CFBundleShortVersionString':'1.09',}
    }

setup(
    app=APP,
    name='Expiry CSV-to-XLSX',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
