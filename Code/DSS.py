# -*- coding: iso-8859-15 -*-

"""

Classe responsável por fazer conexão com o OpenDSS. Seu objeto armazena os atalhos das interfaces"

"""

import ctypes
import win32com.client
from win32com.client import makepy
# from numpy import *
from pylab import *
from comtypes import automation
import os
import sys

author = "Paulo Radatz and Celso Rocha"
version = "01.00.02"
last_update = "09/13/2016"


class DSS(object):

    def __init__(self):
        """ Cria um novo objeto DSS """
        """ Compile OpenDSS model and initialize variables."""

        # These variables provide direct interface into OpenDSS
        sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
        makepy.main()  # ensures early binding and improves speed

        # Create a new instance of the DSS
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Start the DSS
        if self.dssObj.Start(0) == False:
            ### I need to add a window
            print "DSS Failed to Start"

        print "OpenDSS Version: " + self.dssObj.Version

        # Assign a variable to each of the interfaces for easier access
        self.dssText = self.dssObj.Text
        self.dssCircuit = self.dssObj.ActiveCircuit
        self.dssSolution = self.dssCircuit.Solution
        self.dssCktElement = self.dssCircuit.ActiveCktElement
        self.dssBus = self.dssCircuit.ActiveBus
        self.dssMeters = self.dssCircuit.Meters
        self.dssPDElement = self.dssCircuit.PDElements
        self.dssSource = self.dssCircuit.Vsources
        self.dssTransformer = self.dssCircuit.Transformers
        self.dssLines = self.dssCircuit.Lines
        self.dssRegulators = self.dssCircuit.RegControls
        self.dssCapControls = self.dssCircuit.CapControls
