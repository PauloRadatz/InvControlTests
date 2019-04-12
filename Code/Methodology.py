# ##  # -*- coding: iso-8859-15 -*-


import win32com.client
from win32com.client import makepy
from numpy import *
from pylab import *
import os               # for path manipulation and directory operations
import pandas as pd
import numpy as np
import DSS
import Scenario
import random
import Results


class Methodology(object):
    list_false = ["0", "No", "no", "n", "NO", 0, 0.0]
    list_default = ["Default", "default", "d"]
    list_true = ["1", "Yes", "yes", "ny", "YES", 1, 1.0]

    def __init__(self):

        self.dss = DSS.DSS()
        self.resultsObj = Results.Results(self)

    def set_condition(self, condition):

        self.condition = condition
        self.resultsObj.set_condition(condition)

    def get_feeder_demand(self):

        self.compile_dss()
        self.solve_snapshot()

        self.condition.feeder_demand = -1 * self.dss.dssCircuit.TotalPower[0]

    def get_buses(self):

        numNodes_list = []
        node1_list = []
        node2_list = []
        node3_list = []
        busNodes_list = []

        self.compile_dss()
        self.solve_snapshot()

        for bus in self.dss.dssCircuit.AllBusNames:
            busNodes = bus

            self.dss.dssCircuit.SetActiveBus(bus)

            numNodes = self.dss.dssBus.NumNodes
            numNodes_list.append(numNodes)

            nodes = self.dss.dssBus.Nodes

            df_buses = pd.DataFrame()

            if 1 in nodes:
                node1_list.append("yes")
                busNodes = busNodes + ".1"
            else:
                node1_list.append("no")
            if 2 in nodes:
                node2_list.append("yes")
                busNodes = busNodes + ".2"
            else:
                node2_list.append("no")
            if 3 in nodes:
                node3_list.append("yes")
                busNodes = busNodes + ".3"
            else:
                node3_list.append("no")

            busNodes_list.append(busNodes)

        self.condition.df_buses["Bus"] = self.dss.dssCircuit.AllBusNames
        self.condition.df_buses["BusNodes"] = busNodes_list
        self.condition.df_buses["NumNodes"] = numNodes_list
        self.condition.df_buses["Node1"] = node1_list
        self.condition.df_buses["Node2"] = node2_list
        self.condition.df_buses["Node3"] = node3_list

        self.condition.df_buses3 = self.condition.df_buses[(self.condition.df_buses["NumNodes"] == 3) & (self.condition.df_buses["Node1"] == "yes") & (self.condition.df_buses["Node2"] == "yes") & (self.condition.df_buses["Node3"] == "yes")][["Bus", "BusNodes"]]

    def compile_dss(self):

        # Always a good idea to clear the DSS when loading a new circuit
        self.dss.dssObj.ClearAll()

        # Load the given circuit master file into OpenDSS
        line1 = "compile [" + self.condition.dssFileName + "]"
        self.dss.dssText.Command = line1

    def set_pvSystems(self):

        for index, pv in self.condition.df_PVSystems.iterrows():

            busName = pv["PV Bus"]
            bus = pv["BusNodes"]

            self.dss.dssCircuit.SetActiveBus(bus)
            kV = self.dss.dssBus.kVBase * sqrt(3)


            line1 = "New line.PV_{} phases=3 bus1={} bus2=PV_sec_{} switch=yes".format(busName, bus, bus)
            line2 = "New transformer.PV_{} phases=3 windings=2 buses=(PV_sec_{} , PV_ter_{}) conns=(wye, wye) kVs=({},0.48) xhl=5.67 %R=0.4726 kVAs=({},{})".format(busName, bus, bus, kV, pv["kVA"], pv["kVA"])
            line3 = "makebuslist"
            line4 = "setkVBase bus=PV_sec_{} kVLL={}".format(busName, kV)
            line5 = "setkVBase bus=PV_ter_{} kVLL=0.48".format(busName)
            line6 = "New PVSystem.PV_{} phases=3 conn=wye  bus1=PV_ter_{} kV=0.48 kVA={} irradiance=1 Pmpp={} pf=1 %cutin=0.05 %cutout=0.05 VarFollowInverter=yes kvarlimit={}".format(busName, bus, pv["kVA"], pv["Pmpp"], pv["kvarlimit"])

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
        y_curveW = "[1 1 0 0]"
        x_curveW = "[1 1.05 1.1 1.2]"

        line1 = "New XYcurve.generic npts=7 yarray=" + y_curve + " xarray=" + x_curve
        line2 = "New XYcurve.genericW npts=4 yarray=" + y_curveW + " xarray=" + x_curveW

        self.dss.dssText.Command = line1
        self.dss.dssText.Command = line2

        for index, pv in self.condition.df_PVSystems.iterrows():

            smart_function = pv["Smart Functions"]
            busName = pv["PV Bus"]

            if smart_function == "PF" or smart_function == "PF_VW":
                line = "PVSystem.PV_{}.pf={}".format(busName, pv["pf"])

            elif smart_function == "voltvar":
                line = 'New InvControl.{} mode=voltvar voltage_curvex_ref={} vvc_curve1=generic deltaQ_factor={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} VarChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.condition.deltaQ_factor, pv["VV_RefReactivePower"], busName, self.condition.voltageChangeTolerance, self.condition.varChangeTolerance)
                #print("n")

            elif smart_function == "voltwatt":
                line = 'New InvControl.{} mode=voltwatt voltage_curvex_ref={} voltwatt_curve=genericW deltaP_factor={} VoltwattYAxis={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.condition.deltaP_factor, pv["voltwattYAxis"], busName, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance)
                #print("n")
            elif smart_function == "DRC":
                # DRC
                line = "New InvControl.feeder" + busName + \
                       " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.2 EventLog=No PVSystemList=PV_" + busName


            elif smart_function == "VV_VW":
                line = 'New InvControl.{} combimode=VV_VW voltage_curvex_ref={} vvc_curve1=generic voltwatt_curve=genericW deltaP_factor={} deltaQ_factor={} VoltwattYAxis={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={} VarChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.condition.deltaP_factor, self.condition.deltaQ_factor, pv["voltwattYAxis"], pv["VV_RefReactivePower"], busName, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance, self.condition.varChangeTolerance)
                #print("n")

            elif smart_function == "VV_DRC":
                #Volt/var and DRC
                line = "New InvControl.feeder" + busName + \
                       " CombiMode=VV_DRC  voltage_curvex_ref=rated VV_RefReactivePower=varMax_watts vvc_curve1=generic DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=0.2 EventLog=No PVSystemList=PV_" + busName

            self.dss.dssText.Command = line

    def solve_snapshot(self):

        self.dss.dssObj.AllowForms = "false"
        self.dss.dssText.Command = "set mode=snap"
        self.dss.dssText.Command = "set maxcontroliter ={}".format(self.condition.maxControlIter)
        self.dss.dssText.Command = "solve"

    def get_scenario_results(self):

        print self.dss.dssSolution.ControlIterations

        if self.dss.dssSolution.ControlIterations == self.condition.maxControlIter and self.condition.export_scenario_issue in Methodology.list_true:
            self.resultsObj.get_config_issued()

        self.condition.controlIterations.append(self.dss.dssSolution.ControlIterations)

        self.condition.maxVoltage.append(max(self.dss.dssCircuit.AllBusVmagPu))

        self.condition.scenario_feeder_demand.append(-1 * self.dss.dssCircuit.TotalPower[0])


    def get_condition_results(self):

        self.resultsObj.get_scenarios_results()

        self.resultsObj.get_statistics()