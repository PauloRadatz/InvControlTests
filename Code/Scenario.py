# encoding: utf-8

author = "Paulo Radatz and Celso Rocha"
version = "01.00.04"
last_update = "10/13/2017"

import timeit
from pylab import *
import pandas as pd
import os
import itertools
import random


class Settings(object):

    def __init__(self, dssFileName, row, methodologyObj):

        self.methodologyObj = methodologyObj
        self.methodologyObj.set_condition(self)

        self.df_buses = pd.DataFrame()
        self.df_buses1 = pd.DataFrame()
        self.df_buses3 = pd.DataFrame()
        self.df_buses2 = pd.DataFrame()

        self.feeder_demand = nan

        self.scenarioResultsList = []

        self.dssFileName = dssFileName  # OpenDSS file Name
        self.feederName = str(row["Feeder Name"])
        self.conditionID = str(row["Condition ID"])
        self.numberScenarios = int(row["Number of Scenarios"])
        self.percentagePenetrationLevel = row["Penetration Level (%)"]
        self.percentageBuses = row["Buses with PVSystem (%)"]

        if row["DeltaP_factor"] == "default":
            self.deltaP_factor = 1.0
        else:
            self.deltaP_factor = row["DeltaP_factor"]

        if row["DeltaQ_factor"] == "default":
            self.deltaQ_factor = 0.7
        else:
            self.deltaQ_factor = row["DeltaQ_factor"]

        if row["VarChangeTolerance"] == "default":
            self.varChangeTolerance = 0.025
        else:
            self.varChangeTolerance = row["VarChangeTolerance"]

        if row["ActivePChangeTolerance"] == "default":
            self.activePChangeTolerance = 0.01
        else:
            self.activePChangeTolerance = row["ActivePChangeTolerance"]

        if row["VoltageChangeTolerance"] == "default":
            self.voltageChangeTolerance = 0.0001
        else:
            self.voltageChangeTolerance = row["VoltageChangeTolerance"]

        self.simulationMode = row["Simulation Mode"]

        self.methodologyObj.get_buses()

        self.methodologyObj.get_feeder_demand()

        self.penetrationLevel = self.percentagePenetrationLevel / 100.0 * self.feeder_demand

        self.df_buses_selected = self.df_buses3


        # -------------- Directories -------------------#
        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        self.resultsMainPath = self.OpenDSS_folder_path + r"/Results/"
        if not os.path.exists(self.resultsMainPath):
            os.makedirs(self.resultsMainPath)

        # Results Directory
        self.resultsPath = self.resultsMainPath + self.feederName + "_" + str(int(self.penetrationLevel)) + r"/"
        if not os.path.exists(self.resultsPath):
            os.makedirs(self.resultsPath)

        # -------------- End Directories -------------------#

    def process(self, k):

        self.methodologyObj.set_condition(self)

        self.df_PVSystems = pd.DataFrame()
        self.kVA_list = []
        self.kvarlimit_list = []
        self.pctPmpp_list = []
        self.wattPriority_list = []
        self.pfPriority_list = []
        self.mode_list = []

        self.vv_RefReactivePower_list = []
        self.voltage_curvex_ref_list = []
        self.voltwattYAxis_list = []



        # Start counting the time of the simulation of the entire simulation
        start_basescenario_time = timeit.default_timer()



        self.numberBuses = int(self.percentageBuses / 100.0 * len(self.df_buses_selected))
        df_buses_random = self.df_buses_selected.ix[random.sample(self.df_buses_selected.index, self.numberBuses)].reset_index(drop=True)


        for i in range(self.numberBuses):

            self.set_pvsystem_properties()
            self.set_invcontrol_properties()


        self.df_PVSystems["PV Bus"] = df_buses_random["Bus"]
        self.df_PVSystems["BusNodes"] = df_buses_random["BusNodes"]
        self.df_PVSystems["Pmpp"] = 1.0 * self.penetrationLevel / self.numberBuses
        self.df_PVSystems["kVA"] = self.kVA_list
        self.df_PVSystems["kvarlimit"] = self.kvarlimit_list
        self.df_PVSystems["pctPmpp"] = self.pctPmpp_list
        self.df_PVSystems["wattPriority"] = self.wattPriority_list
        self.df_PVSystems["pfPriority"] = self.pfPriority_list
        self.df_PVSystems["Smart Functions"] = self.mode_list
        self.df_PVSystems["VV_RefReactivePower"] = self.vv_RefReactivePower_list
        self.df_PVSystems["voltage_curvex_ref"] = self.voltage_curvex_ref_list
        self.df_PVSystems["voltwattYAxis"] = self.voltwattYAxis_list

        self.runScenario()

        elapsed_scenario = (timeit.default_timer() - start_basescenario_time) / 60

        print "The Total RunTime of scenario " + str(k + 1) + " is: " + str(elapsed_scenario) + " min"



    def runScenario(self):

        # Compile the Master File
        self.methodologyObj.compile_dss()

        # set PVSystems
        self.methodologyObj.set_pvSystems()

        # set smart function
        self.methodologyObj.set_smartfunction()

        # Solve Snap
        self.methodologyObj.solve_snapshot()

        self.methodologyObj.get_results()

    def get_results(self):
        self.methodologyObj.show_results()
        print "here"

    def set_pvsystem_properties(self):

        pmpp = 1.0 * self.penetrationLevel / self.numberBuses

        kVA = [pmpp, 1.1 * pmpp]
        pctPmpp = [60, 100]
        wattPriority = ["yes", "no"]
        pfPriority = ["yes", "no"]

        kVA_value = kVA[randint(0, len(kVA))]

        kvarlimit = [0.44 * kVA_value, kVA_value]

        self.kVA_list.append(kVA_value)
        self.kvarlimit_list.append(kvarlimit[randint(0, len(kvarlimit))])
        self.pctPmpp_list.append(pctPmpp[randint(0, len(pctPmpp))])
        self.wattPriority_list.append(wattPriority[randint(0, len(wattPriority))])
        self.pfPriority_list.append(pfPriority[randint(0, len(pfPriority))])

    def set_invcontrol_properties(self):

        # Inverter smart functions
        if self.simulationMode == "SnapShot":
            mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW']
        else:
            mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW', "DRC", "VV_DRC"]

        vv_RefReactivePower = ['VARAVAL_WATTS', 'VARMAX_WATTS']
        voltage_curvex_ref = ['rated']
        voltwattYAxis = ['PMPPPU', 'PAVAILABLEPU', 'PCTPMPPPU', 'KVARATINGPU']



        self.mode_list.append(mode[randint(0, len(mode))])
        self.vv_RefReactivePower_list.append(vv_RefReactivePower[randint(0, len(vv_RefReactivePower))])
        self.voltage_curvex_ref_list.append(voltage_curvex_ref[randint(0, len(voltage_curvex_ref))])
        self.voltwattYAxis_list.append(voltwattYAxis[randint(0, len(voltwattYAxis))])