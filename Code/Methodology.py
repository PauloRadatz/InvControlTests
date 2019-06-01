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
    list_true = ["1", "Yes", "yes", "y", "YES", 1, 1.0]

    def __init__(self, outputFolder_temp):

        self.dss_communication = None
        self.dss_redirect = True
        self.dss_redirect_only_first = True

        self.outputFolder_temp = outputFolder_temp

        self.dss = DSS.DSS(self.dss_communication)
        self.resultsObj = Results.Results(self)

        self.dss_command = False #True

    def set_condition(self, condition):

        self.condition = condition
        self.resultsObj.set_condition(condition)

    def get_feeder_demand(self):

        self.compile_dss(-1)
        self.solve_snapshot()

        if self.dss_communication == "COM":
            self.condition.feeder_demand = -1 * self.dss.dssCircuit.TotalPower[0]
        else:
            self.condition.feeder_demand = -1 * self.dss.get_Circuit_TotalPower()[0]

    def get_buses(self):

        numNodes_list = []
        node1_list = []
        node2_list = []
        node3_list = []
        busNodes_list = []

        self.compile_dss(-1)
        self.solve_snapshot()

        if self.dss_communication == "COM":
            buses = -1 * self.dss.dssCircuit.TotalPower[0]
        else:
            buses = self.dss.get_Circuit_AllBusNames()

        for bus in buses:
            busNodes = bus

            if self.dss_communication == "COM":
                self.dss.dssCircuit.SetActiveBus(bus)
                numNodes = self.dss.dssBus.NumNodes
                nodes = self.dss.dssBus.Nodes
            else:
                self.dss.set_Circuit_ActiveBus_Name(bus)
                numNodes = self.dss.get_Bus_NumNodes()
                nodes = self.dss.get_Bus_Nodes()

            numNodes_list.append(numNodes)

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

        if self.dss_communication == "COM":
            self.condition.df_buses["Bus"] = self.dss.dssCircuit.AllBusNames
        else:
            self.condition.df_buses["Bus"] = self.dss.get_Circuit_AllBusNames()
        self.condition.df_buses["BusNodes"] = busNodes_list
        self.condition.df_buses["NumNodes"] = numNodes_list
        self.condition.df_buses["Node1"] = node1_list
        self.condition.df_buses["Node2"] = node2_list
        self.condition.df_buses["Node3"] = node3_list

        self.condition.df_buses3 = self.condition.df_buses[(self.condition.df_buses["NumNodes"] == 3) & (self.condition.df_buses["Node1"] == "yes") & (self.condition.df_buses["Node2"] == "yes") & (self.condition.df_buses["Node3"] == "yes")][["Bus", "BusNodes"]]

    def compile_dss(self, scenarioID):
        if self.condition.conditionID == "1":
            self.dss_redirect_only_first = True
        else:
            self.dss_redirect_only_first = False

        # Run test
        self.dss_redirect_only_first = False

        if scenarioID != -1:
            if self.dss_redirect:
                self.dss_pv = self.outputFolder_temp + "/scenario_" + str(scenarioID) +".dss"
                if self.dss_redirect_only_first:
                    self.f = open(self.dss_pv, "w")

        # Always a good idea to clear the DSS when loading a new circuit
        if self.dss_communication == "COM":
            self.dss.dssObj.ClearAll()
        else:
            # Always a good idea to clear the DSS when loading a new circuit
            self.dss.clearAll()
            self.dss.text("clearall")

        # Load the given circuit master file into OpenDSS
        line1 = "compile [" + self.condition.dssFileName + "]"
        line2 = "New XYCurve.Eff npts=4 xarray=[.1 .2 .4 1.0] yarray=[1 0.97 0.96 0.95]"
        line3 = "New XYCurve.FatorPvsT npts=4 xarray=[0 25 75 100] yarray=[1 1 1 1]"

        if self.dss_redirect and scenarioID != -1 and self.dss_redirect_only_first:
            self.f.write("\n" + "!" + line1)
            self.f.write("\n" + line2)
            self.f.write("\n" + line3 + "\n" + "\n")

            if self.dss_communication == "COM":
                self.dss.dssText.Command = line1
            else:
                self.dss.text(line1)

        elif not self.dss_redirect:
            if self.dss_communication == "COM":
                self.dss.dssText.Command = line1
                self.dss.dssText.Command = line2
                self.dss.dssText.Command = line3
            else:
                self.dss.text(line1)
                self.dss.text(line2)
                self.dss.text(line3)
        elif scenarioID == -1:
            if self.dss_communication == "COM":
                self.dss.dssText.Command = line1
                self.dss.dssText.Command = line2
                self.dss.dssText.Command = line3
            else:
                self.dss.text(line1)
                self.dss.text(line2)
                self.dss.text(line3)
        else:
            if self.dss_communication == "COM":
                self.dss.dssText.Command = line1
            else:
                self.dss.text(line1)

            if self.dss_command:
                print line1
                print line2
                print line3

    def set_pvSystems(self):

        if self.dss_redirect_only_first:

            for index, pv in self.condition.df_PVSystems.iterrows():

                busName = pv["PV Bus"]
                bus = pv["BusNodes"]

                if self.dss_communication == "COM":
                    self.dss.dssCircuit.SetActiveBus(bus)
                    kV = self.dss.dssBus.kVBase * sqrt(3)
                else:
                    self.dss.set_Circuit_ActiveBus_Name(bus)
                    kV = self.dss.get_Bus_kVBase() * sqrt(3)

                self.condition.df_PVSystems["kV"][index] = kV

                line1 = "New line.PV_{} phases=3 bus1={} bus2=PV_sec_{} switch=yes".format(busName, bus, bus)
                line2 = "New transformer.PV_{} phases=3 windings=2 buses=(PV_sec_{} , PV_ter_{}) conns=(wye, wye) kVs=({},0.48) xhl=5.67 %R=0.4726 kVAs=({},{})".format(busName, bus, bus, kV, pv["kVA"], pv["kVA"])
                line3 = "makebuslist"
                line4 = "setkVBase bus=PV_sec_{} kVLL={}".format(busName, kV)
                line5 = "setkVBase bus=PV_ter_{} kVLL=0.48".format(busName)
                line6 = "New PVSystem.PV_{} phases=3 conn=wye bus1=PV_ter_{} kV=0.48 kVA={} irradiance=1 Pmpp={} P-TCurve=FatorPvsT EffCurve=Eff %cutin=0.05 %cutout=0.05 VarFollowInverter=yes kvarlimit={} wattpriority={}".format(busName, bus, pv["kVA"], pv["Pmpp"], pv["kvarlimit"], pv["wattPriority"])

                if self.dss_redirect and self.dss_redirect_only_first:
                    self.f.write("\n" + line1)
                    self.f.write("\n" + line2)
                    self.f.write("\n" + line3)
                    self.f.write("\n" + line4)
                    self.f.write("\n" + line5)
                    self.f.write("\n" + line6 + "\n" + "\n")
                elif not self.dss_redirect:
                    if self.dss_communication == "COM":
                        self.dss.dssText.Command = line1
                        self.dss.dssText.Command = line2
                        self.dss.dssText.Command = line3
                        self.dss.dssText.Command = line4
                        self.dss.dssText.Command = line5
                        self.dss.dssText.Command = line6
                    else:
                        self.dss.text(line1)
                        self.dss.text(line2)
                        self.dss.text(line3)
                        self.dss.text(line4)
                        self.dss.text(line5)
                        self.dss.text(line6)

                if self.dss_command:
                    print line1
                    print line2
                    print line3
                    print line4
                    print line5
                    print line6

    def set_smartfunction(self):

        if self.dss_redirect_only_first:
            # Volt-var Curve
            x_curve = "[0.5 0.92 0.95 1.0 1.02 1.05 1.5]"
            y_curve = "[1 1 0 0 0 -1 -1]"

            x_curve = "[0.5 0.95 1.0 1.05 1.5]"
            y_curve = "[1 1 0 -1 -1]"

            # Volt-watt Curve
            y_curveW = "[1 1 0 0]"
            x_curveW = "[1 1.02 1.1 1.2]"

            line1 = "New XYcurve.generic npts=7 yarray=" + y_curve + " xarray=" + x_curve
            line1 = "New XYcurve.generic npts=5 yarray=" + y_curve + " xarray=" + x_curve
            line2 = "New XYcurve.genericW npts=4 yarray=" + y_curveW + " xarray=" + x_curveW

            if self.dss_redirect and self.dss_redirect_only_first:
                self.f.write("\n" + line1)
                self.f.write("\n" + line2 + "\n" + "\n")
            elif not self.dss_redirect:
                if self.dss_communication == "COM":
                    self.dss.dssText.Command = line1
                    self.dss.dssText.Command = line2
                else:
                    self.dss.text(line1)
                    self.dss.text(line2)

            if self.dss_command:
                print line1
                print line2

            for index, pv in self.condition.df_PVSystems.iterrows():

                smart_function = pv["Smart Functions"]
                busName = pv["PV Bus"]

                if smart_function == "PF" or smart_function == "PF_VW":
                    line = "Edit PVSystem.PV_{} pf={} pfpriority={}".format(busName, pv["pf"], pv["pfPriority"])

                elif smart_function == "voltvar":
                    line = 'New InvControl.{} mode=voltvar voltage_curvex_ref={} vvc_curve1=generic deltaQ_factor={} RefReactivePower={} eventlog=yes PVSystemList=PVSystem.PV_{} VoltageChangeTolerance={} VarChangeTolerance={}'\
                        .format(busName, pv["voltage_curvex_ref"], self.condition.deltaQ_factor, pv["RefReactivePower"], busName, self.condition.voltageChangeTolerance, self.condition.varChangeTolerance)
                    #print("n")

                elif smart_function == "voltwatt":
                    line = 'New InvControl.{} mode=voltwatt voltage_curvex_ref={} voltwatt_curve=genericW deltaP_factor={} VoltwattYAxis={} eventlog=yes PVSystemList=PVSystem.PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={}'\
                        .format(busName, pv["voltage_curvex_ref"], self.condition.deltaP_factor, pv["voltwattYAxis"], busName, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance)
                    #print("n")
                elif smart_function == "DRC":
                    # DRC
                    line = "New InvControl.feeder" + busName + \
                           " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.2 EventLog=yes PVSystemList=PVSystem.PV_" + busName


                elif smart_function == "VV_VW":
                    line = 'New InvControl.{} combimode=VV_VW voltage_curvex_ref={} vvc_curve1=generic voltwatt_curve=genericW deltaP_factor={} deltaQ_factor={} VoltwattYAxis={} RefReactivePower={} eventlog=no PVSystemList=PVSystem.PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={} VarChangeTolerance={}'\
                        .format(busName, pv["voltage_curvex_ref"], self.condition.deltaP_factor, self.condition.deltaQ_factor, pv["voltwattYAxis"], pv["RefReactivePower"], busName, self.condition.voltageChangeTolerance, self.condition.activePChangeTolerance, self.condition.varChangeTolerance)
                    #print("n")

                elif smart_function == "VV_DRC":
                    #Volt/var and DRC
                    line = "New InvControl.feeder" + busName + \
                           " CombiMode=VV_DRC  voltage_curvex_ref=rated RefReactivePower=varMax_watts vvc_curve1=generic DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=0.2 EventLog=No PVSystemList=PVSystem.PV_" + busName

                if self.dss_redirect and self.dss_redirect_only_first:
                    self.f.write("\n" + line)
                elif not self.dss_redirect:
                    if self.dss_communication == "COM":
                        self.dss.dssText.Command = line
                    else:
                        self.dss.text(line)

                if self.dss_command:
                    print line

        if self.dss_redirect:
            if self.dss_redirect_only_first:
                self.f.close()
            line_redirect = 'redirect "' + self.dss_pv + '"'
            if self.dss_communication == "COM":
                self.dss.dssText.Command = line_redirect
            else:
                self.dss.text(line_redirect)

    def solve_snapshot(self):

        line1 = "set mode=snap"
        line2 = "set maxcontroliter ={}".format(self.condition.maxControlIter)
        line3 = "set maxiterations=1000"
        line4 = "batchedit invcontrol..* deltaQ_factor={}".format(self.condition.deltaQ_factor)
        line5 = "batchedit invcontrol..* deltaP_factor={}".format(self.condition.deltaP_factor)
        line6 = "solve"

        if self.dss_communication == "COM":
            self.dss.dssObj.AllowForms = "false"
            self.dss.dssText.Command = line1
            self.dss.dssText.Command = line2
            self.dss.dssText.Command = line3
            self.dss.dssText.Command = line4
            self.dss.dssText.Command = line5
            self.dss.dssText.Command = line6

        else:
            self.dss.text(line1)
            self.dss.text(line2)
            self.dss.text(line3)
            self.dss.text(line4)
            self.dss.text(line5)
            self.dss.text(line6)

        if self.dss_command:
            print line1
            print line2
            print line3
            print line4
            print line5
            print line6

    def get_scenario_results(self):

        if self.dss_communication == "COM":
            control_iterations = self.dss.dssSolution.ControlIterations
        else:
            control_iterations = self.dss.get_Solution_ControlIterations()

        print control_iterations

        if control_iterations == self.condition.maxControlIter and self.condition.export_scenario_issue in Methodology.list_true:
            self.resultsObj.get_config_issued()

        self.condition.controlIterations.append(control_iterations)

        if self.dss_communication == "COM":
            self.condition.maxVoltage.append(max(self.dss.dssCircuit.AllBusVmagPu))
            self.condition.scenario_feeder_demand.append(-1 * self.dss.dssCircuit.TotalPower[0])
        else:
            self.condition.maxVoltage.append(max(self.dss.get_Circuit_AllBusVMagPu()))
            self.condition.scenario_feeder_demand.append(-1 * self.dss.get_Circuit_TotalPower()[0])


    def get_condition_results(self):

        self.resultsObj.get_scenarios_results()

        self.resultsObj.get_statistics()