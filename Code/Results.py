# encoding: utf-8



import math
import pandas as pd
import numpy as np
import traceback


class Results:

    def __init__(self, methodologyObj):

        self.condition = None
        self.methodologyObj = methodologyObj

        self.controlIterations = []

        self.df_issue_dic = {}

    def set_condition(self, condition):
        self.condition = condition

    def get_results(self):

        self.df_results = pd.DataFrame()

        self.df_results["Max Control Iterations"] = self.controlIterations

    def get_maxControlIter_statistics(self):

        sr_maxcontroliter = pd.Series(self.df_results["Max Control Iterations"])

        num_issue = sr_maxcontroliter.isin([self.condition.maxControlIter]).sum()

        sr_num_noissue_maxcontroliter = sr_maxcontroliter.where(sr_maxcontroliter.isin([self.condition.maxControlIter]) == False).dropna()

        sr_maxControlIter = sr_num_noissue_maxcontroliter.describe()

        self.maxControlIter_count = sr_maxControlIter["count"]
        self.maxControlIter_mean = sr_maxControlIter["mean"]
        self.maxControlIter_std = sr_maxControlIter["std"]
        self.maxControlIter_min = sr_maxControlIter["min"]
        self.maxControlIter_max = sr_maxControlIter["max"]

    def get_config_issued(self):

        self.df_issue_dic[self.condition.scenarioID] = self.condition.df_PVSystems