# encoding: utf-8



import math
import pandas as pd
import numpy as np
import traceback


class Results:

    def __init__(self, methodologyObj):

        self.condition = None
        self.methodologyObj = methodologyObj

        self.sr_condition_statistics_results_dic = {}



    def set_condition(self, condition):
        self.condition = condition



    def get_condition_results(self):
        self.df_condition_results = pd.DataFrame()

        self.df_condition_results["Max Control Iterations"] = self.condition.controlIterations

    def get_maxControlIter_statistics(self):

        sr_maxcontroliter = pd.Series(self.df_condition_results["Max Control Iterations"])

        num_issue = sr_maxcontroliter.isin([self.condition.maxControlIter]).sum()

        sr_num_noissue_maxcontroliter = sr_maxcontroliter.where(sr_maxcontroliter.isin([self.condition.maxControlIter]) == False).dropna()

        sr_maxControlIter = sr_num_noissue_maxcontroliter.describe()

        self.maxControlIter_count = sr_maxControlIter["count"]
        self.maxControlIter_mean = sr_maxControlIter["mean"]
        self.maxControlIter_std = sr_maxControlIter["std"]
        self.maxControlIter_min = sr_maxControlIter["min"]
        self.maxControlIter_max = sr_maxControlIter["max"]

        self.sr_condition_statistics_results_dic[self.condition.conditionID] = sr_maxControlIter

    def get_config_issued(self):

        self.condition.df_issue_dic[self.condition.scenarioID] = self.condition.df_PVSystems