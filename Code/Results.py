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

    def set_condition(self, condition):
        self.condition = condition

    def get_results(self):

        self.df = pd.DataFrame()

        self.df["Control"] = self.controlIterations