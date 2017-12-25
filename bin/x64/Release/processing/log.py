# -*- coding: utf-8 -*-
"""
Created on Fri Dec  1 09:50:25 2017

@author: Frank
"""

import inspect, logging

def autolog(message):
    "Automatically log the current function details."
    # Get the previous frame in the stack, otherwise it would
    # be this function!!!
    func = inspect.currentframe().f_back.f_code
    # Dump the message + the name of this function to the log.
    logging.debug("%s: %s in %s:%i" % (
        message, 
        func.co_name, 
        func.co_filename, 
        func.co_firstlineno
    ))

def log(message):
    func = inspect.currentframe().f_back.f_code
    print(func.co_firstlineno,message)