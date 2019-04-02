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

        self.df_PVSystems = pd.DataFrame()

        self.scenarioResultsList = []

        self.dssFileName = dssFileName  # OpenDSS file Name
        self.feederName = str(row["Feeder Name"])
        self.conditionID = str(row["Condition ID"])
        self.numberScenarios = int(row["Number of Scenarios"])
        self.penetrationLevel = row["Penetration Level (kW)"]
        self.percentageBuses = row["Buses with PVSystem (%)"]
        self.busesListFile = row["Buses File"]

        self.simulationMode = row["Simulation Mode"]


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

        # Inverter smart functions
        if self.simulationMode == "SnapShot":
            self.smart_functions = {0: 'PF', 1: 'VV', 2: 'VW', 3: 'VV_VW', 4: 'PF_VW'}


    def process(self, k):

        self.methodologyObj.set_condition(self)

        smartFunctionList = []

        # Start counting the time of the simulation of the entire simulation
        start_basescenario_time = timeit.default_timer()

        busesList = pd.read_csv(self.busesListFile, header=None)[0].tolist()

        numberBuses = int(self.percentageBuses / 100.0 * len(busesList))

        busesPV_list = random.sample(busesList, numberBuses)


        for i in range(numberBuses):

            smartFunctionList.append(randint(0, len(self.smart_functions)))

        self.df_PVSystems["PV Buses"] = busesPV_list
        self.df_PVSystems["Smart Functions"] = smartFunctionList
        self.df_PVSystems["size (kW)"] = 1.0 * self.penetrationLevel / numberBuses

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






