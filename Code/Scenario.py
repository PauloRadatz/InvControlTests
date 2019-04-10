# encoding: utf-8



import timeit
from pylab import *
import pandas as pd
import os
import itertools
import random


class Settings(object):
    list_false = ["0", "No", "no", "n", "NO", 0, 0.0]
    list_default = ["Default", "default", "d"] + list_false

    def __init__(self, dssFileName, row, methodologyObj, df_scenario_options):

        self.methodologyObj = methodologyObj
        self.methodologyObj.set_condition(self)

        self.df_scenario_options = df_scenario_options
        self.df_buses = pd.DataFrame()
        self.df_buses1 = pd.DataFrame()
        self.df_buses3 = pd.DataFrame()
        self.df_buses2 = pd.DataFrame()

        self.controlIterations = []
        self.df_issue_dic = {}

        self.feeder_demand = nan

        self.scenarioResultsList = []

        self.dssFileName = dssFileName  # OpenDSS file Name
        self.feederName = str(row["Feeder Name"])
        self.conditionID = str(row["Condition ID"])
        self.numberScenarios = int(row["Number of Scenarios"])
        self.percentagePenetrationLevel = row["Penetration Level (%)"]
        self.percentageBuses = row["Buses with PVSystem (%)"]
        self.maxControlIter = row["Max Control Iterations"]

        self.connections_buses_fixed = row["Scenarios buses fixed"]

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

        self.numberBuses = int(self.percentageBuses / 100.0 * len(self.df_buses_selected))
        self.df_buses_scenarios_fixed = self.df_buses_selected.ix[random.sample(self.df_buses_selected.index, self.numberBuses)].reset_index(drop=True)


        self.dic_buses_scenarios = {}

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

    def process(self, k, df_buses_fixed, base, fixed):

        self.methodologyObj.set_condition(self)

        self.scenarioID = k + 1

        self.df_PVSystems = pd.DataFrame()
        self.kVA_list = []
        self.kvarlimit_list = []
        self.pctPmpp_list = []
        self.wattPriority_list = []
        self.pfPriority_list = []
        self.mode_list = []
        self.pf_list = []

        self.vv_RefReactivePower_list = []
        self.voltage_curvex_ref_list = []
        self.voltwattYAxis_list = []



        # Start counting the time of the simulation of the entire simulation
        start_basescenario_time = timeit.default_timer()


        for i in range(self.numberBuses):

            self.set_pvsystem_properties()
            self.set_invcontrol_properties()

        if base == "yes" or fixed == "no":

            if self.df_scenario_options["Fixed"]["ScenariosBusesFixed"] in Settings.list_false:
                df_buses = self.df_buses_selected.ix[random.sample(self.df_buses_selected.index, self.numberBuses)].reset_index(drop=True)
            else:
                df_buses = self.df_buses_scenarios_fixed

            self.dic_buses_scenarios[k] = df_buses

            if k == (self.numberScenarios - 1):
                self.df_buses_scenarios = pd.concat(self.dic_buses_scenarios, axis=1)

        elif fixed == "yes":
            df_buses = df_buses_fixed[k]



        self.df_PVSystems["PV Bus"] = df_buses["Bus"]
        self.df_PVSystems["BusNodes"] = df_buses["BusNodes"]
        self.df_PVSystems["Pmpp"] = 1.0 * self.penetrationLevel / self.numberBuses
        self.df_PVSystems["kVA"] = self.kVA_list
        self.df_PVSystems["pf"] = self.pf_list
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



        print "The Total RunTime of scenario " + str(self.scenarioID) + " is: " + str(elapsed_scenario) + " min"



    def runScenario(self):

        # Compile the Master File
        self.methodologyObj.compile_dss()

        # set PVSystems
        self.methodologyObj.set_pvSystems()

        # set smart function
        self.methodologyObj.set_smartfunction()

        # Solve Snap
        self.methodologyObj.solve_snapshot()

        self.methodologyObj.get_scenario_results()

    def get_results(self):
        self.methodologyObj.get_condition_results()
        print "here"

    def set_pvsystem_properties(self):

        pmpp = 1.0 * self.penetrationLevel / self.numberBuses

        # kVA property
        if self.df_scenario_options["Fixed"]["kVA"] in Settings.list_false:
            kVA = [pmpp, 1.1 * pmpp]
        elif self.df_scenario_options["Value"]["kVA"] in Settings.list_default:
            kVA = [pmpp]
        else:
            kVA = []

        # pctPmpp property
        if self.df_scenario_options["Fixed"]["pctPmpp"] in Settings.list_false:
            pctPmpp = [60.0, 100.0]
        elif self.df_scenario_options["Value"]["pctPmpp"] in Settings.list_default:
            pctPmpp = [100.0]
        else:
            pctPmpp = [float(self.df_scenario_options["Value"]["pctPmpp"])]

        # WattPriority property
        if self.df_scenario_options["Fixed"]["wattPriority"] in Settings.list_false:
            wattPriority = ["yes", "no"]
        elif self.df_scenario_options["Value"]["wattPriority"] in Settings.list_default:
            wattPriority = ["no"]
        else:
            wattPriority = [self.df_scenario_options["Value"]["WattPriority"]]

        # PFPriority property
        if self.df_scenario_options["Fixed"]["pfPriority"] in Settings.list_false:
            pfPriority = ["yes", "no"]
        elif self.df_scenario_options["Value"]["pfPriority"] in Settings.list_default:
            pfPriority = ["yes"]
        else:
            pfPriority = [self.df_scenario_options["Value"]["pfPriority"]]

        # pf property
        if self.df_scenario_options["Fixed"]["pf"] in Settings.list_false:
            pf = ["-0.98", "0.98"]
        elif self.df_scenario_options["Value"]["pf"] in Settings.list_default:
            pf = ["-0.98"]
        else:
            pf = [self.df_scenario_options["Value"]["pf"]]

        kVA_value = kVA[randint(0, len(kVA))]

        # kvarlimit property
        if self.df_scenario_options["Fixed"]["kvarlimit"] in Settings.list_false:
            kvarlimit = [0.44 * kVA_value, kVA_value]
        elif self.df_scenario_options["Value"]["kvarlimit"] in Settings.list_default:
            kvarlimit = [kVA_value]
        else:
            kvarlimit = []

        # Populates lists
        self.kVA_list.append(kVA_value)
        self.kvarlimit_list.append(kvarlimit[randint(0, len(kvarlimit))])
        self.pctPmpp_list.append(pctPmpp[randint(0, len(pctPmpp))])
        self.wattPriority_list.append(wattPriority[randint(0, len(wattPriority))])
        self.pfPriority_list.append(pfPriority[randint(0, len(pfPriority))])
        self.pf_list.append(pf[randint(0, len(pf))])

    def set_invcontrol_properties(self):

        # Inverter smart functions
        if self.simulationMode == "SnapShot":
            mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW']
        else:
            mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW', "DRC", "VV_DRC"]

        # vv_RefReactivePower property
        if self.df_scenario_options["Fixed"]["vv_RefReactivePower"] in Settings.list_false:
            vv_RefReactivePower = ['VARAVAL_WATTS', 'VARMAX_WATTS']
        elif self.df_scenario_options["Value"]["vv_RefReactivePower"] in Settings.list_default:
            vv_RefReactivePower = ['VARMAX_WATTS']
        else:
            vv_RefReactivePower = []

        # voltage_curvex_ref property
        if self.df_scenario_options["Fixed"]["voltage_curvex_ref"] in Settings.list_false:
            voltage_curvex_ref = ['rated']
        elif self.df_scenario_options["Value"]["voltage_curvex_ref"] in Settings.list_default:
            voltage_curvex_ref = ['rated']
        else:
            voltage_curvex_ref = []

        # voltwattYAxis property
        if self.df_scenario_options["Fixed"]["voltwattYAxis"] in Settings.list_false:
            voltwattYAxis = ['PMPPPU', 'PAVAILABLEPU', 'PCTPMPPPU', 'KVARATINGPU']
        elif self.df_scenario_options["Value"]["voltwattYAxis"] in Settings.list_default:
            voltwattYAxis = ['PMPPPU']
        else:
            voltwattYAxis = []


        self.mode_list.append(mode[randint(0, len(mode))])
        self.vv_RefReactivePower_list.append(vv_RefReactivePower[randint(0, len(vv_RefReactivePower))])
        self.voltage_curvex_ref_list.append(voltage_curvex_ref[randint(0, len(voltage_curvex_ref))])
        self.voltwattYAxis_list.append(voltwattYAxis[randint(0, len(voltwattYAxis))])