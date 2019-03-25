###  # -*- coding: iso-8859-15 -*-

author = "Paulo Radatz and Celso Rocha"
version = "01.00.02"
last_update = "09/13/2016"

import win32com.client
from win32com.client import makepy
from numpy import *
from pylab import *
import os


def comInterfaceConnected():
    """ Compile OpenDSS model and initialize variables."""

    # These variables provide direct interface into OpenDSS
    sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
    makepy.main()  # ensures early binding and improves speed

    # Create a new instance of the DSS
    dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

    # Start the DSS
    if dssObj.Start(0) == False:
        ### I need to add a window
        print "DSS Failed to Start"

    print "OpenDSS Version: " + dssObj.Version

    return dssObj