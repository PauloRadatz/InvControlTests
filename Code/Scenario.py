# encoding: utf-8

import timeit
from pylab import *
import pandas as pd
import random


class Settings(object):
    list_false = ["0", "No", "no", "n", "NO", 0, 0.0]
    list_default = ["Default", "default", "d"]
    list_true = ["1", "Yes", "yes", "ny", "YES", 1, 1.0]

    def __init__(self, dssFileName, row, methodologyObj, df_scenario_options):

        # Reads inputs
        self.dssFileName = dssFileName
        self.methodologyObj = methodologyObj
        self.df_scenario_options = df_scenario_options

        # Actives this condition object into the methodology object
        self.methodologyObj.set_condition(self)

        # ----- Variables initialization -----#
        # Buses dataframes
        self.df_buses = pd.DataFrame()
        self.df_buses1 = pd.DataFrame()
        self.df_buses3 = pd.DataFrame()
        self.df_buses2 = pd.DataFrame()

        # ----- Scenario Results ---- #
        # Control iterations for each scenario
        self.controlIterations = []
        # Max voltage
        self.maxVoltage = []
        # Scenario simulation time
        self.scenario_simulation_time = []
        # Feeder Demand
        self.scenario_feeder_demand = []
        # dataframe that stores the scenarios informations of those which have a Max control iteration issue
        self.df_issue_dic = {}

        # Variable that stores the max feeder demand. It is calculated in Methodogy.get_feeder_demand method
        self.feeder_demand = nan

        ###########################
        self.scenarioResultsList = []

        # Reads condition definitions
        self.feederName = str(row["Feeder Name"])
        self.conditionID = str(row["Condition ID"])
        self.numberScenarios = int(row["Number of Scenarios"])
        #self.percentagePenetrationLevel = row["Penetration Level (%)"]
        self.penetrationLevel_read = row["Penetration Level (kW)"]
        self.percentageBuses = row["Buses with PVSystem (%)"]
        self.maxControlIter = row["Max Control Iterations"]
        self.simulationMode = row["Simulation Mode"]
        self.export_scenario_issue = row["Export Scenario Issue"]

        # Control loop convergence parameters
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

        # Populates buses dataframes
        self.methodologyObj.get_buses()

        # Gets max feeder demand and stores it in self.feeder_demand
        self.methodologyObj.get_feeder_demand()

        # defines the penetration level in kW
        #self.penetrationLevel = self.percentagePenetrationLevel / 100.0 * self.feeder_demand
        self.penetrationLevel = self.penetrationLevel_read

        # FOR NOW IT HAS ONLY THE OPTION FOR THREE-PHASE BUSES.
        # HERE WE NEED TO INCLUDE OTHER OPTIONS THAT WILL BE SET IN THE CONDITION INPUT
        self.df_buses_selected = self.df_buses3

        # Defines the number of buses with PVSystem
        self.numberBuses = int(self.percentageBuses / 100.0 * len(self.df_buses_selected))

        # Creates a buses dataframe in case the user wants to have all scenarios with the same buses.
        # Check ScenariosBusesFixed
        self.df_buses_scenarios_fixed = self.df_buses_selected.ix[random.sample(self.df_buses_selected.index, self.numberBuses)].reset_index(drop=True)

        # Dictionary used to store the scenarios of the base connection to be used in the others ones
        self.scenarios_fixed_dic = {}

    def set_methodology(self, metodology):

        self.methodologyObj = metodology

    def process(self, k, df_scenarios_fixed, base, fixed):

        # Actives this condition object into the methodology object
        self.methodologyObj.set_condition(self)

        # Scenario identification
        self.scenarioID = k

        # PVSystem lists
        self.kVA_list = []
        self.kvarlimit_list = []
        self.pctPmpp_list = []
        self.wattPriority_list = []
        self.pfPriority_list = []
        self.mode_list = []
        self.pf_list = []

        # InvControl Lists
        self.RefReactivePower_list = []
        self.voltage_curvex_ref_list = []
        self.voltwattYAxis_list = []

        # Scenario information
        self.df_PVSystems = pd.DataFrame()

        # For each PV deployment, the parameters are set randomly
        for i in range(self.numberBuses):
            self.set_pvsystem_properties()
            self.set_invcontrol_properties()

        # If this is a base (1st) condition of the option to have the same scenarios through the conditions are not set
        if base in Settings.list_true or fixed in Settings.list_false:

            # If the need to have the same buses through all scenarios are not set
            if self.df_scenario_options["Fixed"]["ScenariosBusesFixed"] in Settings.list_false:
                df_buses = self.df_buses_selected.ix[random.sample(self.df_buses_selected.index, self.numberBuses)].reset_index(drop=True)
            # Needs to have all scenarios with the same buses
            else:
                df_buses = self.df_buses_scenarios_fixed

            # Populates the scenario information
            self.df_PVSystems["PV Bus"] = df_buses["Bus"]
            self.df_PVSystems["BusNodes"] = df_buses["BusNodes"]
            self.df_PVSystems["kV"] = 0.0
            self.df_PVSystems["Pmpp"] = 1.0 * self.penetrationLevel / self.numberBuses
            self.df_PVSystems["kVA"] = self.kVA_list
            self.df_PVSystems["pf"] = self.pf_list
            self.df_PVSystems["kvarlimit"] = self.kvarlimit_list
            self.df_PVSystems["pctPmpp"] = self.pctPmpp_list
            self.df_PVSystems["wattPriority"] = self.wattPriority_list
            self.df_PVSystems["pfPriority"] = self.pfPriority_list
            self.df_PVSystems["Smart Functions"] = self.mode_list
            self.df_PVSystems["RefReactivePower"] = self.RefReactivePower_list
            self.df_PVSystems["voltage_curvex_ref"] = self.voltage_curvex_ref_list
            self.df_PVSystems["voltwattYAxis"] = self.voltwattYAxis_list

            # Stores it in case this scenarios is used in another condition
            self.scenarios_fixed_dic[k] = self.df_PVSystems

            # In the last scenarios, the dataframe with all scenarios informations is created
            if k == (self.numberScenarios - 1):
                self.df_scenarios_fixed = pd.concat(self.scenarios_fixed_dic, axis=1)

        # Conditions use the same scenarios from the base one
        elif fixed in Settings.list_true:
            self.df_PVSystems = df_scenarios_fixed[k]

        # Runs this scenario
        self.runScenario()
        self.get_scenario_results()

    def runScenario(self):

        self.methodologyObj.compile_dss(self.scenarioID)
        self.methodologyObj.set_pvSystems()
        self.methodologyObj.set_smartfunction()

        # Start counting the time of the scenario simulation time
        start_time = timeit.default_timer()

        # Solve Snap
        self.methodologyObj.solve_snapshot()

        scenario_total_time = (timeit.default_timer() - start_time)
        self.scenario_simulation_time.append(scenario_total_time * 1000)

        print "The Total RunTime of scenario " + str(self.scenarioID) + " is: " + str(scenario_total_time) + " s"

    def get_scenario_results(self):
        self.methodologyObj.get_scenario_results()

    def get_results(self):
        self.methodologyObj.get_condition_results()

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
            if self.df_scenario_options["Fixed"]["Smart Functions"] in Settings.list_false:
                mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW']
            elif self.df_scenario_options["Fixed"]["Smart Functions"] == "Q":
                mode = ['PF', 'voltvar']
            elif self.df_scenario_options["Fixed"]["Smart Functions"] == "P":
                mode = ['voltwatt']
            else:
                mode = ['voltvar']
        else:
            mode = ['PF', 'voltvar', 'voltwatt', 'VV_VW', 'PF_VW', "DRC", "VV_DRC"]

        # RefReactivePower property
        if self.df_scenario_options["Fixed"]["RefReactivePower"] in Settings.list_false:
            RefReactivePower = ['VARAVAL', 'VARMAX']
        elif self.df_scenario_options["Value"]["RefReactivePower"] in Settings.list_default:
            RefReactivePower = ['VARAVAL']
        else:
            RefReactivePower = ['VARAVAL']

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
        self.RefReactivePower_list.append(RefReactivePower[randint(0, len(RefReactivePower))])
        self.voltage_curvex_ref_list.append(voltage_curvex_ref[randint(0, len(voltage_curvex_ref))])
        self.voltwattYAxis_list.append(voltwattYAxis[randint(0, len(voltwattYAxis))])