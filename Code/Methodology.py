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
import DSS
import Scenario

class Methodology(object):

    #------------------------------------------------------------------------------------------------------------------#
    def __init__(self):

        # OpenDSS Object
        self.dss = DSS.DSS()

    def set_condition(self, condition):
        self.condition = condition

    def compile_dss(self):

        # Always a good idea to clear the DSS when loading a new circuit
        self.dss.dssObj.ClearAll()

        # Load the given circuit master file into OpenDSS
        line1 = "compile [" + self.condition.dssFileName + "]"
        self.dss.dssText.Command = line1

    def set_pvSystems(self):

        for index, pv in self.condition.df_PVSystems.iterrows():

            kWPmpp = float(pv["Pmpp"])
            pvkVA = 1.1 * kWPmpp
            bus = pv["PV Buses"]

            self.dss.dssCircuit.SetActiveBus(bus)
            kV = self.dss.dssBus.kVBase * sqrt(3)

            if kWPmpp > 0:
                line1 = "New line.PV_{} phases=3 bus1={} bus2=PV_sec_{} switch=yes".format(bus, bus, bus)
                line2 = "New transformer.PV_{} phases=3 windings=2 buses=(PV_sec_{} , PV_ter_{}) conns=(wye, wye) kVs=({},0.48) xhl=5.67 %R=0.4726 kVAs=({},{})".format(bus, bus, bus, kV, pv["kVA"], pv["kVA"])
                line3 = "makebuslist"
                line4 = "setkVBase bus=PV_sec_{} kVLL={}".format(bus, kV)
                line5 = "setkVBase bus=PV_ter_{} kVLL=0.48".format(bus)
                line6 = "New PVSystem.PV_{} phases=3 conn=wye  bus1=PV_ter_{} kV=0.48 kVA={} irradiance=1 Pmpp={} pf=1 %cutin=0.05 %cutout=0.05 VarFollowInverter=yes kvarlimit={}".format(bus, bus, pv["kVA"], pv["Pmpp"], pv["kvarlimit"])

                self.dss.dssText.Command = line1
                self.dss.dssText.Command = line2
                self.dss.dssText.Command = line3
                self.dss.dssText.Command = line4
                self.dss.dssText.Command = line5
                self.dss.dssText.Command = line6


    def set_smartfunction(self):

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

        for index, pv in self.condition.df_PVSystems.iterrows():

            smart_function = pv["Smart Functions"]
            bus = pv["PV Buses"]

            if smart_function == "PF" or smart_function == "PF_VW":
                line = "PVSystem.PV_" + bus + ".pf=-0.97"

            elif smart_function == "voltvar":
                line = 'New InvControl.{} mode=voltvar voltage_curvex_ref={} vvc_curve1=generic deltaQ_factor={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} VarChangeTolerance={}'\
                    .format(bus, pv["voltage_curvex_ref"], self.condition.deltaQ_factor, pv["VV_RefReactivePower"], bus, self.condition.voltageChangeTolerance, self.condition.varChangeTolerance)
                #print("n")

            elif smart_function == "voltwatt":
                line = 'New InvControl.{} mode=voltwatt voltage_curvex_ref={} voltwatt_curve=genericW deltaP_factor={} VoltwattYAxis={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={}'\
                    .format(bus, pv["voltage_curvex_ref"], self.condition.deltaP_factor, pv["voltwattYAxis"], bus, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance)
                #print("n")
            elif smart_function == "DRC":
                # DRC
                line = "New InvControl.feeder" + bus + \
                       " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.2 EventLog=No PVSystemList=PV_" + bus


            elif smart_function == "VV_VW":
                line = 'New InvControl.{} combimode=VV_VW voltage_curvex_ref={} vvc_curve1=generic voltwatt_curve=genericW deltaP_factor={} deltaQ_factor={} VoltwattYAxis={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={} VarChangeTolerance={}'\
                    .format(bus, pv["voltage_curvex_ref"], self.condition.deltaP_factor, self.condition.deltaQ_factor, pv["voltwattYAxis"], pv["VV_RefReactivePower"], bus, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance, self.condition.varChangeTolerance)
                #print("n")

            elif smart_function == "VV_DRC":
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