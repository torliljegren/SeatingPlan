"""
Usage:
    python setup.py py2app
"""
from glob import glob
from setuptools import setup

APP = ['__init__.py']
DATA_FILES = [
    ('', glob('*.png')),
    ('', ['about.txt'])
]
OPTIONS = {
    "iconfile": "lpico.png",
    "plist": {"CFBundleShortVersionString": "1.0"}
}

setup(
    app=APP,
    name="Sittplaceraren",
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app']
)
