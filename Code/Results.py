# encoding: utf-8

import pandas as pd

class Results:

    def __init__(self, methodologyObj):

        self.condition = None
        self.methodologyObj = methodologyObj

        self.sr_condition_statistics_results_dic = {}
        self.df_scenarios_results_dic = {}



    def set_condition(self, condition):
        self.condition = condition



    def get_scenarios_results(self):
        self.df_scenarios_results = pd.DataFrame()

       # self.df_scenarios_results["Scenario ID"] = range(self.condition.scenarioID)
        self.df_scenarios_results["# Control Iterations"] = self.condition.controlIterations
        self.df_scenarios_results["Max Voltage (pu)"] = self.condition.maxVoltage
        self.df_scenarios_results["Simulation Time (ms)"] = self.condition.scenario_simulation_time
        self.df_scenarios_results["Max feeder Demand without PVs (kW)"] = self.condition.feeder_demand
        self.df_scenarios_results["Max feeder Demand with PVs (kW)"] = self.condition.scenario_feeder_demand
        self.df_scenarios_results["Penetration Level (kW)"] = self.condition.penetrationLevel

        self.df_scenarios_results_without_issue = self.df_scenarios_results[self.df_scenarios_results["# Control Iterations"] != self.condition.maxControlIter]
        self.df_scenarios_results_with_issue = self.df_scenarios_results[self.df_scenarios_results["# Control Iterations"] == self.condition.maxControlIter]

        self.df_scenarios_results_dic[self.condition.conditionID] = self.df_scenarios_results

    def get_scenarios_results_qsts(self):
        self.df_scenarios_results = pd.DataFrame()

       # self.df_scenarios_results["Scenario ID"] = range(self.condition.scenarioID)
        self.df_scenarios_results["# Control Iterations"] = self.condition.controlIterations

        self.df_scenarios_results_without_issue = self.df_scenarios_results[self.df_scenarios_results["# Control Iterations"] != self.condition.maxControlIter]
        self.df_scenarios_results_with_issue = self.df_scenarios_results[self.df_scenarios_results["# Control Iterations"] == self.condition.maxControlIter]

        self.df_scenarios_results_dic[self.condition.conditionID] = self.df_scenarios_results

    def get_statistics(self):

        self.sr_condition_results = pd.Series()

        sr_maxControlIter = pd.Series(self.df_scenarios_results_without_issue["# Control Iterations"]).describe()
        sr_simulationTime = pd.Series(self.df_scenarios_results_without_issue["Simulation Time (ms)"]).describe()

        self.sr_condition_results["Condition ID"] = self.condition.conditionID
        self.sr_condition_results["# Scenarios"] = self.condition.numberScenarios
        self.sr_condition_results["Scenarios without Issue (%)"] = sr_maxControlIter["count"] * 100.0 / self.condition.numberScenarios
        self.sr_condition_results["mean - # Control Iteration"] = sr_maxControlIter["mean"]
        self.sr_condition_results["std - # Control Iteration"] = sr_maxControlIter["std"]
        self.sr_condition_results["min - # Control Iteration"] = sr_maxControlIter["min"]
        self.sr_condition_results["max - # Control Iteration"] = sr_maxControlIter["max"]
        self.sr_condition_results["mean - Simulation Time (ms)"] = sr_simulationTime["mean"]
        self.sr_condition_results["std - Simulation Time (ms)"] = sr_simulationTime["std"]
        self.sr_condition_results["min - Simulation Time (ms)"] = sr_simulationTime["min"]
        self.sr_condition_results["max - Simulation Time (ms)"] = sr_simulationTime["max"]
        self.sr_condition_results["DeltaP_factor"] = self.condition.deltaP_factor
        self.sr_condition_results["DeltaQ_factor"] = self.condition.deltaQ_factor


        self.sr_condition_statistics_results_dic[self.condition.conditionID] = self.sr_condition_results

    def get_statistics_qsts(self):

        self.sr_condition_results = pd.Series()

        sr_maxControlIter = pd.Series(self.df_scenarios_results_without_issue["# Control Iterations"]).describe()

        self.sr_condition_results["Condition ID"] = self.condition.conditionID
        self.sr_condition_results["# Scenarios"] = self.condition.numberScenarios
        self.sr_condition_results["Scenarios without Issue (%)"] = sr_maxControlIter["count"] * 100.0 / self.condition.numberScenarios
        self.sr_condition_results["mean - # Control Iteration"] = sr_maxControlIter["mean"]
        self.sr_condition_results["std - # Control Iteration"] = sr_maxControlIter["std"]
        self.sr_condition_results["min - # Control Iteration"] = sr_maxControlIter["min"]
        self.sr_condition_results["max - # Control Iteration"] = sr_maxControlIter["max"]
        self.sr_condition_results["DeltaP_factor"] = self.condition.deltaP_factor
        self.sr_condition_results["DeltaQ_factor"] = self.condition.deltaQ_factor


        self.sr_condition_statistics_results_dic[self.condition.conditionID] = self.sr_condition_results

    def get_config_issued(self):

        self.condition.df_issue_dic[self.condition.scenarioID] = self.condition.df_PVSystems