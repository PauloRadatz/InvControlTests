# encoding: utf-8

author = "Paulo Radatz and Celso Rocha"
version = "01.00.04"
last_update = "10/13/2017"

import ReadResults
import ControlOpenDSS
import timeit
from pylab import *
import pandas as pd
import os
import itertools


class Settings(object):

    def __init__(self, dssObj, dssFileName, row):

        self.scenarioResultsList = []

        self.dssObj = dssObj  # OpenDSS Object
        self.dssFileName = dssFileName  # OpenDSS file Name
        self.feederName = str(row["Feeder Name"])
        self.connectionID = str(row["Connection ID"])
        self.numberPV = row["Number of PV Sites"]  # Number of PVs in the system
        self.loadinglevel = row["Feeder Load Conditions"]
        self.pvGenCurve = row["Solar Profile"]
        self.pvLocationNames = row["PV Location Names"].split("_")
        self.pvLocations = [str(i) for i in row["PV Locations"].split(" ")]
        self.pvSizes_kW = [float(i) for i in row["PV sizes"].split("_")]
        self.pvLocationsPF = [str(i) for i in row["Power Factor"].split("_")]  # Power factor used in the smart inverter function Fixed PF level 2
        self.numRegulators = row["Number of Regulators"]  # Includes OLTC
        self.numCapacitors = row["Number of Capacitors"]
        self.reactivePowerCapacitors = row["Reactive Power (kvar) Capacitors"]

        self.penetrationLevel = sum(self.pvSizes_kW)

        # -------------- Directories -------------------#
        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        self.resultsMainPath = self.OpenDSS_folder_path + r"/Results_TimeSeries_HC/"
        if not os.path.exists(self.resultsMainPath):
            os.makedirs(self.resultsMainPath)

        # Results Directory
        self.resultsPath = self.resultsMainPath + str(int(self.penetrationLevel)) + "kW_" + str(self.numberPV) + "PVs_" + str(self.pvGenCurve) + "_" + str(self.loadinglevel) + r"/"
        if not os.path.exists(self.resultsPath):
            os.makedirs(self.resultsPath)

        # -------------- End Directories -------------------#

        # ------ Limits Power Quality Metric Values --------#

        self.max_srpdi = 0.25 * self.reactivePowerCapacitors  # Considering that capacitors stay on for 6h max
        max_srpdi = 0.0
        for i in range(self.numberPV):
            max_srpdi = max_srpdi + 0.32 * ((self.pvSizes_kW[i] / float(self.pvLocationsPF[i])) ** 2 - self.pvSizes_kW[i] ** 2) ** 0.5

        if max_srpdi > self.max_srpdi:
            self.max_srpdi = max_srpdi

        self.max_savmvi = 0.0333
        self.max_savfi = 3.0 / 100
        self.max_savui = 0.0678
        self.max_scdoi_vr = 10000.0 / 365  # 27.4
        self.max_scdoi_cap = 4.0
        self.max_seli = 9.0 / 100

        # ----- End Limits Power Quality Metric Values ------#

        # Inverter smart functions
        #self.smart_functions = {0: 'PF', 1: 'VV', 2: 'VW', 3: 'DRC', 4: 'VV_VW', 5: 'VV_DRC'}
        self.smart_functions = {0: 'PF', 1: 'VV', 2: 'VW'}

        # Volt-var Curve
        self.x_curve = "[0.5 0.92 0.98 1.0 1.02 1.08 1.5]"
        self.y_curve = "[1 1 0 0 0 -1 -1]"

        # Volt-watt Curve
        self.y_curveW = "[1 1 0]"
        self.x_curveW = "[1 1.05 1.1]"

        # To calculate VVI, the base case PCC's monitors path should be stored
        self.baseCasePCCMonitorsPaths = []

        # Scenarios
        # this list stores the name of the combination functions
        self.scenarioName = []
        # It is a list of lists. Each list stores 3 integers that relate to the smart inverter possible functions
        self.scenarios = list(itertools.product(range(len(self.smart_functions)), repeat=int(self.numberPV)))

    def process(self):
        """
        This function is the main code.
        """
        # Start counting the time of the simulation of the entire simulation
        start_basescenario_time = timeit.default_timer()

        # Set DataPath and set caseName defined in OpenDSS. To calculate VVI, the base case results path should be stored
        self.baseCaseResultsPath = self.resultsPath + "Base"
        self.caseName = "Base_Base_Base"
        self.folder_DI = self.resultsPath + r"Base/Base_Base_Base"

        # Run the Base Case scenario. The base case has all PVSystems operating with unity power factor
        self.runPVBaseCase()
        self.scenarioName.append(["Base", "Base", "Base"])

        # PCC monitors path
        for i in range(self.numberPV):
            self.baseCasePCCMonitorsPaths.append(self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_terminalV_" + str(i) + "_1.csv")

        # Feeder Head Monitors Path
        feederHead_monVoltPath = self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_FeederHeadV_1.csv"
        feederHead_monPowerPath = self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_FeederHeadP_1.csv"

        # Feeder End Monitor Path
        feederEnd_monVoltPath = self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_FeederEndV_1.csv"

        # Regulators Monitors Path
        regulators_monTapsPaths = []
        if self.regulators_monTaps is not None:
            for i in range(len(self.regulators_monTaps)):
                regulators_monTapsPaths.append(self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_Regulator_" + str(i) + "_1.csv")

        # CapControls Monitors Path
        capControls_monCapsPaths = []
        if self.capControls_monCaps is not None:
            for i in range(len(self.capControls_monCaps)):
                capControls_monCapsPaths.append(self.baseCaseResultsPath + r"/" + self.caseName + "_Mon_SwitchedCap_" + str(i) + "_1.csv")

        # Get Base Case Results
        # DI results
        self.getDIFileResults()

        # Feeder Head Voltage
        resultsVoltFeederHead = self.getVoltOneMon(feederHead_monVoltPath, self.feederHead_Vbase)
        self.maxFeederHeadVoltage = resultsVoltFeederHead[0]
        self.minFeederHeadVoltage = resultsVoltFeederHead[1]
        self.meanFeederHeadVoltage = resultsVoltFeederHead[2]

        # Feeder End Voltage
        resultsVoltFeederEnd = self.getVoltOneMon(feederEnd_monVoltPath, self.feederEnd_Vbase)
        self.maxFeederEndVoltage = resultsVoltFeederEnd[0]
        self.minFeederEndVoltage = resultsVoltFeederEnd[1]
        self.meanFeederEndVoltage = resultsVoltFeederEnd[2]

        # PCCs Voltage
        pccVoltage = [480 / sqrt(3), 480 / sqrt(3), 480 / sqrt(3)]
        resultsVoltPCCs = self.getVoltMons(self.baseCasePCCMonitorsPaths, pccVoltage)
        self.maxPCCsVoltage = resultsVoltPCCs[0]
        self.minPCCsVoltage = resultsVoltPCCs[1]
        self.meanPCCsVoltage = resultsVoltPCCs[2]

        if self.regulators_monTaps is not None:
            # Taps Operations
            self.tapOperations = self.getTapOperations(regulators_monTapsPaths)
        else:
            self.tapOperations = 0.0

        if self.capControls_monCaps is not None:
            # Caps Operations
            self.capOperations = self.getCapOperations(capControls_monCapsPaths)
        else:
            self.capOperations = 0.0

        # Power Factor at Feeder Head
        self.maxkWFeederHead, self.minkWFeederHead, self.maxkvarFeederHead, self.minkvarFeederHead, self.minimumFeederHeadPowerFactor = self.getFeederHeadInformation(feederHead_monPowerPath)

        # VVI results are 1 for the base case
        self.vviAtInverterTerminals = [1] * len(self.pvLocations)
        self.maxControlIterationIssues = str(0)

        # This is a internal variable which helps us to eliminate scenarios with convergence issues
        # We can set it manually
        scenariosWconvergenceIssue = []

        # ----------- Power Quality Score -------------#

        # Calculates voltages index
        self.calc_voltage_index(self.pvBaseCase.sr_voltage_average_dic, self.pvBaseCase.sr_voltage_unbalance_dic)

        # Calculates index
        self.calc_index()

        # Gets the scenario results
        self.resultsList()

        elapsed_scenario = (timeit.default_timer() - start_basescenario_time) / 60

        print "The Total RunTime of scenario " + self.caseName + " is: " + str(elapsed_scenario) + " min"

        # Run the scenarios
        for i in range(len(self.scenarios)):
            # Start counting the time of the simulation of the entire simulation
            start_scenario_time = timeit.default_timer()

            if self.scenarios[i] not in scenariosWconvergenceIssue:

                self.casePCCMonitorsPaths = []

                # Set DataPath and set caseName defined in OpenDSS
                self.caseName = str(self.scenarios[i]).replace("(", "").replace(")", "")\
                    .replace(" ", "").replace(",", "_").replace("0", "PF").replace("1", "VV").replace("2", "VW")\
                    .replace("3", "DRC").replace("4", "VV+VW").replace("5", "VV+DRC")

                self.caseResultsPath = self.resultsPath + self.caseName
                self.folder_DI = self.caseResultsPath + r"/" + self.caseName

                # Run a scenario with smart inverter functions
                self.runPVCase(self.scenarios[i])
                self.scenarioName.append([self.smart_functions[self.scenarios[i][0]], self.smart_functions[self.scenarios[i][1]], self.smart_functions[self.scenarios[i][2]]])

                # PCC monitors path
                for j in range(self.numberPV):
                    self.casePCCMonitorsPaths.append(self.caseResultsPath + r"/" + self.caseName + "_Mon_terminalV_" + str(j) + "_1.csv")

                # Feeder Head Monitors Path
                feederHead_monVoltPath = self.caseResultsPath + r"/" + self.caseName + "_Mon_FeederHeadV_1.csv"
                feederHead_monPowerPath = self.caseResultsPath + r"/" + self.caseName + "_Mon_FeederHeadP_1.csv"

                # Feeder End Monitor Path
                feederEnd_monVoltPath = self.caseResultsPath + r"/" + self.caseName + "_Mon_FeederEndV_1.csv"

                # Regulators Monitors Path
                regulators_monTapsPaths = []
                if self.regulators_monTaps is not None:
                    for j in range(len(self.regulators_monTaps)):
                        regulators_monTapsPaths.append(self.caseResultsPath + r"/" + self.caseName + "_Mon_Regulator_" + str(j) + "_1.csv")

                # CapControls Monitors Path
                capControls_monCapsPaths = []
                if self.capControls_monCaps is not None:
                    for j in range(len(self.capControls_monCaps)):
                        capControls_monCapsPaths.append(self.caseResultsPath + r"/" + self.caseName + "_Mon_SwitchedCap_" + str(j) + "_1.csv")


                # Get Scenario Results

                # DI results
                self.getDIFileResults()

                # Feeder Head Voltage
                resultsVoltFeederHead = self.getVoltOneMon(feederHead_monVoltPath, self.feederHead_Vbase)
                self.maxFeederHeadVoltage = resultsVoltFeederHead[0]
                self.minFeederHeadVoltage = resultsVoltFeederHead[1]
                self.meanFeederHeadVoltage = resultsVoltFeederHead[2]

                # Feeder End Voltage
                resultsVoltFeederEnd = self.getVoltOneMon(feederEnd_monVoltPath, self.feederEnd_Vbase)
                self.maxFeederEndVoltage = resultsVoltFeederEnd[0]
                self.minFeederEndVoltage = resultsVoltFeederEnd[1]
                self.meanFeederEndVoltage = resultsVoltFeederEnd[2]

                # PCCs Voltage
                resultsVoltPCCs = self.getVoltMons(self.casePCCMonitorsPaths, pccVoltage)
                self.maxPCCsVoltage = resultsVoltPCCs[0]
                self.minPCCsVoltage = resultsVoltPCCs[1]
                self.meanPCCsVoltage = resultsVoltPCCs[2]

                if self.regulators_monTaps is not None:
                    # Taps Operations
                    self.tapOperations = self.getTapOperations(regulators_monTapsPaths)
                else:
                    self.tapOperations = 0.0

                if self.capControls_monCaps is not None:
                    # Caps Operations
                    self.capOperations = self.getCapOperations(capControls_monCapsPaths)
                else:
                    self.capOperations = 0.0

                # Power Factor at Feeder Head and min powers
                self.maxkWFeederHead, self.minkWFeederHead, self.maxkvarFeederHead, self.minkvarFeederHead, self.minimumFeederHeadPowerFactor = self.getFeederHeadInformation(feederHead_monPowerPath)

                # VVI results
                self.vviAtInverterTerminals = self.getVVI(self.baseCasePCCMonitorsPaths, self.casePCCMonitorsPaths)

                # ----------- Power Quality Score -------------#

                # Calculates voltages index
                self.calc_voltage_index(self.pvCase.sr_voltage_average_dic, self.pvCase.sr_voltage_unbalance_dic)

                # Calculates index
                self.calc_index()

                # Gets the scenario results
                self.resultsList()

                elapsed_scenario = (timeit.default_timer() - start_scenario_time) / 60

                print "The Total RunTime of scenario " + self.scenarioName[i+1][0].replace("_", "+") + "_" + self.scenarioName[i+1][1].replace("_", "+") \
                      + "_" + self.scenarioName[i+1][2].replace("_", "+") + " is: " + str(elapsed_scenario) + " min"

    def calc_index(self):

        self.scdoi_vr_value = float(self.tapOperations) / self.numRegulators
        self.scdoi_cap_value = float(self.capOperations) / self.numCapacitors

        self.srpdi_value = float(self.kvarhFeederHead) / 24.0
        self.seli_value = float(self.feederLosses) / float(self.feederConsumption)


    def calc_voltage_index(self, sr_voltage_average_dic, sr_voltage_unbalance_dic):

        vmaxpu = 1.05
        vminpu = 0.95

        df_voltage_average = pd.DataFrame(sr_voltage_average_dic)
        df_voltage_unbalance = pd.DataFrame(sr_voltage_unbalance_dic)

        # DataFrame Shapes
        t = df_voltage_average.shape[1]
        n = df_voltage_average.shape[0]

        # (1) System Average Voltage Magnitude Violation Index (SAVMVI)
        self.savmvi_value = ((df_voltage_average[df_voltage_average > vmaxpu] - vmaxpu).fillna(0) + (vminpu - df_voltage_average[df_voltage_average < vminpu]).fillna(0)).sum(axis=1).sum() / (t * n)

        # (2) System Average Voltage Fluctuation index (SAVFI)
        self.savfi_value = ((df_voltage_average.diff(axis=1) ** 2) ** 0.5).sum(axis=1).sum() / (t * n)

        # (3) System Average Voltage Unbalance index (SAVUI)
        self.savui_value = df_voltage_unbalance.sum(axis=1).sum() / (t * n)


    def runPVBaseCase(self):
        """
        This function runs the base case scenario.
        The base case scenario has all PVSystems operating with unity power factor
        """

        # Compile the Master File
        self.pvBaseCase = ControlOpenDSS.DSS(self.dssObj, self.dssFileName, self.OpenDSS_folder_path, self.baseCaseResultsPath, self.connectionID, self.caseName)

        # Add Monitor at the FeederHead, FeederEnd, Regulators and Capacitors
        self.feederHead_monVolt, self.feederHead_monPower, self.feederHead_Vbase = self.pvBaseCase.setFeederHeadMonitors()
        self.regulators_monTaps = self.pvBaseCase.setRegulatorsMonitors()
        self.capControls_monCaps = self.pvBaseCase.setCapacitorsMonitors()
        self.feederEnd_monVolt, self.feederEnd_Vbase = self.pvBaseCase.setFeederEndMonitors() ###problemas aqui

        self.pvBaseCase.redirect_curves(self.loadinglevel, self.pvGenCurve)
        self.vBases_BaseCase = self.pvBaseCase.pvSystems(self.pvLocations, self.pvSizes_kW)

        self.pvBaseCase.compileMonitors(self.feederHead_monVolt, self.feederHead_monPower, self.regulators_monTaps, self.capControls_monCaps, self.feederEnd_monVolt)

        # Solve a time-series simulation
        self.pvBaseCase.solve_yearly()

        # Export the monitors
        self.pvBaseCase.export_Monitors(self.pvLocations, self.feederHead_monVolt, self.feederHead_monPower, self.regulators_monTaps, self.capControls_monCaps, self.feederEnd_monVolt)

    def runPVCase(self, scenarioName):
        """
        This function runs scenarios. Scenarios are defined with different smart invert functions
        """

        # Compile the Master File
        self.pvCase = ControlOpenDSS.DSS(self.dssObj, self.dssFileName, self.OpenDSS_folder_path, self.caseResultsPath, self.connectionID, self.caseName)

        self.pvCase.redirect_curves(self.loadinglevel, self.pvGenCurve)
        self.vBases_Scenario = self.pvCase.pvSystems(self.pvLocations, self.pvSizes_kW)
        self.pvCase.inverter(self.y_curve, self.x_curve, self.y_curveW, self.x_curveW, scenarioName, self.pvLocations, self.pvLocationsPF)

        self.pvBaseCase.compileMonitors(self.feederHead_monVolt, self.feederHead_monPower, self.regulators_monTaps, self.capControls_monCaps, self.feederEnd_monVolt)

        # Solve a time-series simulation
        self.pvCase.solve_yearly()

        # Export the monitors
        self.pvCase.export_Monitors(self.pvLocations, self.feederHead_monVolt, self.feederHead_monPower, self.regulators_monTaps, self.capControls_monCaps, self.feederEnd_monVolt)

    def getDIFileResults(self):
        """
        This function gets the results read from the DI files.
        Notice that those DI files are always placed at self.folder_DI directory
        """

        # Results selected from the DI_VoltExceptions.csv file
        # The ReadResults file has a function called readVolt_Exceptions that read information from a DI file
        volt_Exceptions_ListResults = ReadResults.readVolt_Exceptions(self.folder_DI)
        self.max_Voltage = volt_Exceptions_ListResults[0]
        self.max_LV_Voltage = volt_Exceptions_ListResults[6]
        self.max_Voltage_Overvoltages = volt_Exceptions_ListResults[1]
        self.max_LV_Voltage_Overvoltages = volt_Exceptions_ListResults[7]
        self.max_Voltage_Bus = str(volt_Exceptions_ListResults[2])
        self.max_LV_Voltage_Bus = volt_Exceptions_ListResults[8]

        self.min_Voltage = volt_Exceptions_ListResults[3]
        self.min_LV_Voltage = volt_Exceptions_ListResults[9]
        self.min_Voltage_Undervoltages = volt_Exceptions_ListResults[4]
        self.min_LV_Voltage_Undervoltages = volt_Exceptions_ListResults[10]
        self.min_Voltage_Bus = str(volt_Exceptions_ListResults[5])
        self.min_LV_Voltage_Bus = volt_Exceptions_ListResults[11]

        self.timeAboveANSIMaxVoltage = volt_Exceptions_ListResults[12]
        self.timeBelowANSIMinVoltage = volt_Exceptions_ListResults[13]

        # Results selected from the DI_Totals.csv file
        totals = ReadResults.readTotals(self.folder_DI)
        self.kWhFeederHead = float(totals[0])
        self.kvarhFeederHead = float(totals[1])
        #self.maxkWFeederHead = float(totals[2])
        #self.maxkvarFeederHead = float(totals[3])
        self.feederConsumption = float(totals[4])
        self.feederLosses = float(totals[6])

    def getVVI(self, baseVmonitorPaths, scenarioVmonitorPaths):

        vviLocationList = []

        for i in range(len(scenarioVmonitorPaths)):
            scenarioVVI = ReadResults.read_mons_Volt_to_Calculate_VVI(baseVmonitorPaths[i], scenarioVmonitorPaths[i])
            vviLocationList.append(scenarioVVI[3])

            # Check if the scenarios has a maximum control iteration issue
            if i == 0:
                self.maxControlIterationIssues = scenarioVVI[4]

        return vviLocationList

    def getOneCapOperations(self, scenarioCAPmonitorPath):

        return ReadResults.operations_counter(scenarioCAPmonitorPath, None, None)

    def getCapOperations(self, scenarioCAPmonitorPaths):

        capLocationList = []

        for i in range(len(scenarioCAPmonitorPaths)):
            capLocationList.append(self.getOneTapOperations(scenarioCAPmonitorPaths[i]))

        return sum(capLocationList)

    def getOneTapOperations(self, scenarioTAPmonitorPath):

        return ReadResults.operations_counter(scenarioTAPmonitorPath, None, None)

    def getTapOperations(self, scenarioTAPmonitorPaths):

        tapLocationList = []

        for i in range(len(scenarioTAPmonitorPaths)):
            tapLocationList.append(self.getOneTapOperations(scenarioTAPmonitorPaths[i]))

        return sum(tapLocationList)

    def getVoltOneMon(self, scenarioVmonitorPath, vBase):

        voltMonitorResults = ReadResults.read_mon_Volt(scenarioVmonitorPath, vBase)

        vMax = voltMonitorResults[4]
        vMin = voltMonitorResults[5]
        vMean = voltMonitorResults[6]

        return [vMax, vMin, vMean]

    def getVoltMons(self, scenarioVmonitorPaths, vBase):

        vMaxLocationList = []
        vMinLocationList = []
        vMeanLocationList = []

        for i in range(len(scenarioVmonitorPaths)):
            voltMonitorResults = self.getVoltOneMon(scenarioVmonitorPaths[i], vBase[i])
            vMaxLocationList.append(voltMonitorResults[0])
            vMinLocationList.append(voltMonitorResults[1])
            vMeanLocationList.append(voltMonitorResults[2])

        return [vMaxLocationList, vMinLocationList, vMeanLocationList]

    def getFeederHeadInformation(self, scenarioPmonitorPath):

        results = ReadResults.read_mon_Power_feederHead(scenarioPmonitorPath)
        ptMax = results[11]
        ptMin = results[12]
        qtMax = results[13]
        qtMin = results[14]
        pfMin = results[15]

        return ptMax, ptMin, qtMax, qtMin, pfMin

    def resultsList(self):
        """
        This function stores all results that will be sent to MasterConnections.py
        """

        columns =["Near",
                "Middle",
                "End",
                "maxControlIterationIssues",
                "Max Feeder MV Voltage (pu)",
                "Overvoltages at Max Feeder Voltage",
                "max Feeder Voltage Bus Name",
                "Max Feeder LV Voltage (pu)",
                "Overvoltages LV at Max Feeder Voltage",
                "max Feeder Voltage LV Bus Name",
                "Min Feeder MV Voltage (pu)",
                "Undervoltages at Min Feeder Voltage",
                "min Feeder Voltage Bus Name",
                "Min Feeder LV Voltage (pu)",
                "Underoltages LV at Min Feeder Voltage",
                "min Feeder Voltage LV Bus Name",
                "Max kW Feeder Head (kW)",
                "Min kW Feeder Head (kW)",
                "Max kvar Feeder Head (kvar)",
                "Min kvar Feeder Head (kvar)",
                "kWh Feeder Head (kWh)",
                "kvarh Feeder Head (kvarh)",
                "minimum Feeder Head Power Factor",
                "tap Operations",
                "cap Operations",
                "Feeder Losses (kWh)",
                "Feeder Consumption (kWh)",
                "Time Above ANSI Max Voltage (hour)",
                "Time Below ANSI Min Voltage (hour)",
                "max Feeder Head Voltage (pu)",
                "min Feeder Head Voltage (pu)",
                "mean Feeder Head Voltage (pu)",
                "max Feeder End Voltage (pu)",
                "min Feeder End Voltage (pu)",
                "mean Feeder End Voltage (pu)",
                "max PCCs Voltage Near (pu)",
                "max PCCs Voltage Middle (pu)",
                "max PCCs Voltage End (pu)",
                "min PCCs Voltage Near (pu)",
                "min PCCs Voltage Middle (pu)",
                "min PCCs Voltage End (pu)",
                "mean PCCs Voltage Near (pu)",
                "mean PCCs Voltage Middle (pu)",
                "mean PCCs Voltage End (pu)",
                "VVI at Inverter Terminals Near",
                "VVI at Inverter Terminals Middle",
                "VVI at Inverter Terminals End",
                "SAVMI",
                "SAVFI",
                "SAVUI",
                "SCDOI_vr",
                "SCDOI_cap",
                "SRPDI",
                "SELI"]

        list = [self.caseName.split("_")[0],
                self.caseName.split("_")[1],
                self.caseName.split("_")[2],
                self.maxControlIterationIssues,
                self.max_Voltage,
                self.max_Voltage_Overvoltages,
                self.max_Voltage_Bus,
                self.max_LV_Voltage,
                self.max_LV_Voltage_Overvoltages,
                self.max_LV_Voltage_Bus,
                self.min_Voltage,
                self.min_Voltage_Undervoltages,
                self.min_Voltage_Bus,
                self.min_LV_Voltage,
                self.min_LV_Voltage_Undervoltages,
                self.min_LV_Voltage_Bus,
                self.maxkWFeederHead,
                self.minkWFeederHead,
                self.maxkvarFeederHead,
                self.minkvarFeederHead,
                self.kWhFeederHead,
                self.kvarhFeederHead,
                self.minimumFeederHeadPowerFactor,
                self.tapOperations,
                self.capOperations,
                self.feederLosses,
                self.feederConsumption,
                self.timeAboveANSIMaxVoltage,
                self.timeBelowANSIMinVoltage,
                self.maxFeederHeadVoltage,
                self.minFeederHeadVoltage,
                self.meanFeederHeadVoltage,
                self.maxFeederEndVoltage,
                self.minFeederEndVoltage,
                self.meanFeederEndVoltage,
                self.maxPCCsVoltage[0],
                self.maxPCCsVoltage[1],
                self.maxPCCsVoltage[2],
                self.minPCCsVoltage[0],
                self.minPCCsVoltage[1],
                self.minPCCsVoltage[2],
                self.meanPCCsVoltage[0],
                self.meanPCCsVoltage[1],
                self.meanPCCsVoltage[2],
                self.vviAtInverterTerminals[0],
                self.vviAtInverterTerminals[1],
                self.vviAtInverterTerminals[2],
                self.savmvi_value,
                self.savfi_value,
                self.savui_value,
                self.scdoi_vr_value,
                self.scdoi_cap_value,
                self.srpdi_value,
                self.seli_value]

        self.scenarioResultsList.append(pd.Series(list, index=columns))
