# ##  # -*- coding: iso-8859-15 -*-
author = "Paulo Radatz"
version = "01.00.02"
last_update = "09/13/2016"

import win32com.client
from win32com.client import makepy
from numpy import *
from pylab import *
import os               # for path manipulation and directory operations
import pandas as pd
import numpy as np

class DSS(object):

    #------------------------------------------------------------------------------------------------------------------#
    def __init__(self, dssObj, dssFileName):

        # OpenDSS Object
        self.dss = dssObj

        # Always a good idea to clear the DSS when loading a new circuit
        self.dss.dssObj.ClearAll()

        # Load the given circuit master file into OpenDSS
        line1 = "compile [" + dssFileName + "]"
        self.dss.dssText.Command = line1

    def set_pvSystems(self, df_PVSystems):

        for index, pv in df_PVSystems.iterrows():

            kWPmpp = float(pv["size (kW)"])
            pvkVA = 1.1 * kWPmpp
            bus = pv["PV Buses"]

            self.dss.dssCircuit.SetActiveBus(bus)
            kV = self.dss.dssBus.kVBase * sqrt(3)

            if kWPmpp > 0:
                line1 = "New line.PV_" + bus + " phases=3 bus1=" + bus + " bus2=PV_sec_" + bus + " switch=yes"
                line2 = "New transformer.PV_" + bus + " phases=3 windings=2 buses=(PV_sec_" + bus + ", PV_ter_" + bus + ") conns=(wye, wye) kVs=(" + str(kV) + ",0.48) xhl=5.67 %R=0.4726 kVAs=(" + str(pvkVA) + "," + str(pvkVA) + ")"          ##kVAs=(4500, 4500)"
                line3 = "makebuslist"
                line4 = "setkVBase bus=PV_sec_" + bus + " kVLL=" + str(kV)
                line5 = "setkVBase bus=PV_ter_" + bus + " kVLL=0.48"
                line6 = "New PVSystem.PV_" + bus + " phases=3 conn=wye  bus1=PV_ter_" + bus + " kV=0.48 kVA=" + str(pvkVA) + " irradiance=1 Pmpp=" + str(kWPmpp) + " pf=1 %cutin=0.05 %cutout=0.05 VarFollowInverter=yes kvarlimit=" + str(pvkVA*0.44)
                line7 = "New monitor.terminalV_" + bus + " element=transformer.PV_" + bus + " terminal=2 mode=0"
                line8 = "New monitor.terminalP_" + bus + " element=transformer.PV_" + bus + " terminal=2 mode=1 ppolar=no"

                self.dss.dssText.Command = line1
                self.dss.dssText.Command = line2
                self.dss.dssText.Command = line3
                self.dss.dssText.Command = line4
                self.dss.dssText.Command = line5
                self.dss.dssText.Command = line6
                self.dss.dssText.Command = line7
                self.dss.dssText.Command = line8

    def set_smartfunction(self, df_PVSystems):

        # Volt-var Curve
        x_curve = "[0.5 0.92 0.95 1.0 1.02 1.05 1.5]"
        y_curve = "[1 1 0 0 0 -1 -1]"

        # Volt-watt Curve
        y_curveW = "[1 1 0]"
        x_curveW = "[1 1.05 1.1]"

        line1 = "New XYcurve.generic npts=7 yarray=" + y_curve + " xarray=" + x_curve
        line2 = "New XYcurve.genericW npts=4 yarray=" + y_curveW + " xarray=" + x_curveW

        self.dss.dssText.Command = line1
        self.dss.dssText.Command = line2

        for index, pv in df_PVSystems.iterrows():

            smart_function = pv["Smart Functions"]
            bus = pv["PV Buses"]

            if smart_function == 0:
                line = "PVSystem.PV_" + bus + ".pf=-0.97"

            elif smart_function == 1:
                # Volt/var
                line = "New InvControl.feeder" + bus + \
                       " mode=VOLTVAR voltage_curvex_ref=rated vvc_curve1=generic deltaQ_factor=0.2 VV_RefReactivePower=VARMAX_WATTS eventlog=no PVSystemList=PV_" + bus


            elif smart_function == 2:
                # Volt/watt
                line = "New InvControl.feeder" + bus + \
                       " mode=VOLTWATT voltage_curvex_ref=rated voltwatt_curve=genericW  VoltwattYAxis=PMPPPU DeltaP_factor=0.2 eventlog=no PVSystemList=PV_" + bus


            elif smart_function == 3:
                # DRC
                line = "New InvControl.feeder" + bus + \
                       " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.2 EventLog=No PVSystemList=PV_" + bus


            elif smart_function == 4:
                # Volt/watt and Volt/var
                line = "New InvControl.feeder" + bus + \
                       " CombiMode=VV_VW voltage_curvex_ref=rated voltwatt_curve=genericW  VoltwattYAxis=PMPPPU vvc_curve1=generic deltaQ_factor=0.2 VV_RefReactivePower=VARAVAL_WATT eventlog=no PVSystemList=PV_" + bus

            elif smart_function == 5:
                #Volt/var and DRC
                line = "New InvControl.feeder" + bus + \
                       " CombiMode=VV_DRC  voltage_curvex_ref=rated VV_RefReactivePower=varMax_watts vvc_curve1=generic DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=0.2 EventLog=No PVSystemList=PV_" + bus

            self.dss.dssText.Command = line

    def solve_snapshot(self):

        self.dss.dssText.Command = "set maxcontroli = 2000"
        self.dss.dssText.Command = "solve"

        print(self.dss.dssSolution.Converged)
        print(self.dss.dssSolution.ControlIterations)


 #------------------------------------------------------------------------------------------------------------------#