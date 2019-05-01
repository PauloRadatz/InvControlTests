# encoding: utf-8

import os
import Tkinter as tk
import tkFileDialog
import tkMessageBox
import datetime
import pandas as pd
import sqlite3
from sqlalchemy.types import String
import random


class Main(object):
    ConditionList = []
    list_false = ["0", "No", "no", "n", "NO", 0, 0.0]
    list_default = ["Default", "default", "d"]
    list_true = ["1", "Yes", "yes", "ny", "YES", 1, 1.0]

    @classmethod
    def ask_files(cls):

        # Way to close the window that appears with no reason
        root = tk.Tk()
        root.withdraw()

        print "\nPlease select the scenario (*.csv file)."
        #conditionConfiguration_file = tkFileDialog.askopenfilename()
        scenario_file = r"D:\test-1\Condition_1\scenario_9.csv"

        outputFolder = os.path.dirname(scenario_file)

        try:
            config = cls(scenario_file, outputFolder)
            print("Finished" + str(config))
            tkMessageBox.showinfo("", "Finished")
        except Exception as e:
            raise
            print("Incomplete")
            tkMessageBox.showwarning("", "Incomplete")
        return

    def __init__(self, scenario, outputFolder):

        self.deltaQ_factor = 0.2
        self.deltaP_factor = 0.2

        self.voltageChangeTolerance = 0.001
        self.activePChangeTolerance = 0.01
        self.varChangeTolerance = 0.025


        # Read the conditions file
        self.df_scenario = pd.read_csv(scenario, engine="python")

        self.f = open(scenario.split(".")[0] + ".dss", "w")
        self.f.write("!Scenario")

        self.set_pvSystems()
        self.set_smartfunction()

        self.f.close()

    def set_pvSystems(self):

        for index, pv in self.df_scenario.iterrows():
            busName = pv["PV Bus"]
            bus = pv["BusNodes"]

            #kV = pv["kV"]
            kV=10

            line1 = "New line.PV_{} phases=3 bus1={} bus2=PV_sec_{} switch=yes".format(busName, bus, bus)
            line2 = "New transformer.PV_{} phases=3 windings=2 buses=(PV_sec_{} , PV_ter_{}) conns=(wye, wye) kVs=({},0.48) xhl=5.67 %R=0.4726 kVAs=({},{})".format(
                busName, bus, bus, kV, pv["kVA"], pv["kVA"])
            line3 = "makebuslist"
            line4 = "setkVBase bus=PV_sec_{} kVLL={}".format(busName, kV)
            line5 = "setkVBase bus=PV_ter_{} kVLL=0.48".format(busName)
            line6 = "New PVSystem.PV_{} phases=3 conn=wye  bus1=PV_ter_{} kV=0.48 kVA={} irradiance=1 Pmpp={} pf=1 %cutin=0.05 %cutout=0.05 VarFollowInverter=yes kvarlimit={}".format(
                busName, bus, pv["kVA"], pv["Pmpp"], pv["kvarlimit"])

            self.f.write("\n" + line1)
            self.f.write("\n" + line2)
            self.f.write("\n" + line3)
            self.f.write("\n" + line4)
            self.f.write("\n" + line5)
            self.f.write("\n" + line6 + "\n" + "\n")

    def set_smartfunction(self):

        # Volt-var Curve
        x_curve = "[0.5 0.92 0.95 1.0 1.02 1.05 1.5]"
        y_curve = "[1 1 0 0 0 -1 -1]"

        # Volt-watt Curve
        y_curveW = "[1 1 0 0]"
        x_curveW = "[1 1.05 1.1 1.2]"

        line1 = "New XYcurve.generic npts=7 yarray=" + y_curve + " xarray=" + x_curve
        line2 = "New XYcurve.genericW npts=4 yarray=" + y_curveW + " xarray=" + x_curveW

        self.f.write("\n" + line1)
        self.f.write("\n" + line2 + "\n" + "\n")


        for index, pv in self.df_scenario.iterrows():

            smart_function = pv["Smart Functions"]
            busName = pv["PV Bus"]

            if smart_function == "PF" or smart_function == "PF_VW":
                line = "PVSystem.PV_{}.pf={}".format(busName, pv["pf"])

            elif smart_function == "voltvar":
                line = 'New InvControl.{} mode=voltvar voltage_curvex_ref={} vvc_curve1=generic deltaQ_factor={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} VarChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.deltaQ_factor, pv["VV_RefReactivePower"], busName, self.voltageChangeTolerance, self.varChangeTolerance)
                #print("n")

            elif smart_function == "voltwatt":
                line = 'New InvControl.{} mode=voltwatt voltage_curvex_ref={} voltwatt_curve=genericW deltaP_factor={} VoltwattYAxis={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.deltaP_factor, pv["voltwattYAxis"], busName, self.voltageChangeTolerance, self.activePChangeTolerance)
                #print("n")
            elif smart_function == "DRC":
                # DRC
                line = "New InvControl.feeder" + busName + \
                       " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.2 EventLog=No PVSystemList=PV_" + busName


            elif smart_function == "VV_VW":
                line = 'New InvControl.{} combimode=VV_VW voltage_curvex_ref={} vvc_curve1=generic voltwatt_curve=genericW deltaP_factor={} deltaQ_factor={} VoltwattYAxis={} VV_RefReactivePower={} eventlog=no PVSystemList=PV_{} VoltageChangeTolerance={} ActivePChangeTolerance={} VarChangeTolerance={}'\
                    .format(busName, pv["voltage_curvex_ref"], self.deltaP_factor, self.deltaQ_factor, pv["voltwattYAxis"], pv["VV_RefReactivePower"], busName, self.voltageChangeTolerance, self.activePChangeTolerance, self.varChangeTolerance)
                #print("n")

            elif smart_function == "VV_DRC":
                #Volt/var and DRC
                line = "New InvControl.feeder" + busName + \
                       " CombiMode=VV_DRC  voltage_curvex_ref=rated VV_RefReactivePower=varMax_watts vvc_curve1=generic DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=0.2 EventLog=No PVSystemList=PV_" + busName

            self.f.write("\n" + line)

if __name__ == '__main__':
    Main.ask_files()