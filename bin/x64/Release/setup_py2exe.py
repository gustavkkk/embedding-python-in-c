# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 16:11:37 2017

@author: Frank
"""
from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')
setup(
		options = {'py2exe': {'bundle_files': 1}},
		windows = [{'script': "extract.py"}],
		zipfile = None,
)