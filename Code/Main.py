# encoding: utf-8

import Scenario
import DSS
import Methodology
import os
import csv
import timeit
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

        print "\nPlease select a OpenDSS Feeder Model."
        #dssFileName = tkFileDialog.askopenfilename()
        dssFileName = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\IEEE13Nodeckt.dss"

        print "\nPlease select the definitions of the Conditions (*.csv file)."
        #conditionConfiguration_file = tkFileDialog.askopenfilename()
        conditionConfiguration_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input.csv"
        conditionConfiguration_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input_13.csv"

        print "\nPlease select the Output Folder."
        #outputFolder = tkFileDialog.askdirectory()
        outputFolder = r"D:\test-1"

        try:
            config = cls(conditionConfiguration_file, dssFileName, outputFolder)
            print("Finished" + str(config))
            tkMessageBox.showinfo("", "Finished")
        except Exception as e:
            raise
            print("Incomplete")
            tkMessageBox.showwarning("", "Incomplete")
        return

    def __init__(self, fileName, dssFileName, outputFolder):

        # Data and time of the start of simulation
        startSimulation = str(datetime.datetime.now())

        # Start timing of the whole simulation
        start_time = timeit.default_timer()

        print("Simulation starts: " + str(startSimulation))

        # Scenario Options file
        df_scenario_options = pd.read_csv(os.path.dirname(fileName) + r"\Scenario_options.csv", engine="python").set_index("Option")
        # Read the conditions file
        df_conditions = pd.read_csv(fileName, engine="python")

        # Creates a methodology object
        methodologyObj = Methodology.Methodology()

        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        # DataFrame that will store scenarios of the first (base) condition
        df_scenarios_fixed = ""
        base = "yes"

        # Creates condition Objects
        for index, row in df_conditions.iterrows():
            print(u"Scenario ID " + str(row[1]) + u" Created")
            # Each connection is an object of the Settings class
            conditionObject = Scenario.Settings(dssFileName, row, methodologyObj, df_scenario_options)
            Main.ConditionList.append(conditionObject)

        # Runs Connection Objects
        for condition in Main.ConditionList:
            print(u"\n-----------------Running Condition ID " + str(condition.conditionID) + u"--------------------")

            # Sets condition outputfolder
            condition_outputFolder = outputFolder + "/Condition_" + str(condition.conditionID)
            if not os.path.exists(condition_outputFolder):
                os.makedirs(condition_outputFolder)

            # Creates the database that stores the configurations of the scenarios with Max control iterations issue
            sqlite_file = condition_outputFolder + r"\Scenarios_issue.db"

            # Start timing of the connection ID simulation
            start_time_connection = timeit.default_timer()

            # ------------ Main Process ------------#
            for i in range(condition.numberScenarios):
                # Runs each scenario
                condition.process(i, df_scenarios_fixed, base, df_scenario_options["Fixed"]["ScenariosFixed"])

            # Saves the scenarios informations to be used in the following conditions.
            # So, df_scenarios_fixed is created through the first condition
            if base == "yes":
                df_scenarios_fixed = condition.df_scenarios_fixed
                base = "no"

            # Simulation time of the condition
            condition_simulation_time = str(timeit.default_timer() - start_time_connection)
            print(u"\nSimulation time Condition ID " + str(condition.conditionID) + " = " + condition_simulation_time + u" s")

            # Gets results
            condition.get_results()

            # Makes a database Condition
            conn = sqlite3.connect(sqlite_file)
            print(u"\n-----------------Reading Condition ID " + str(condition.conditionID) + u" Results--------------------")

            df_scenarios_results = pd.DataFrame().from_dict(methodologyObj.resultsObj.df_scenarios_results)

            df_scenarios_results.to_csv(condition_outputFolder + r"\scenarios_results.csv", index=False)


            if condition.export_scenario_issue in Main.list_true:
                for issue in condition.df_issue_dic:
                    condition.df_issue_dic[issue].to_sql("scenario_" + str(issue), conn, if_exists="replace")
                    condition.df_issue_dic[issue].to_csv(condition_outputFolder + "/scenario_" + str(issue) + ".csv")

            conn.close()

        df_condition_results = pd.DataFrame().from_dict(methodologyObj.resultsObj.sr_condition_statistics_results_dic)

        df_condition_results.T.to_csv(outputFolder + r"\conditions_results.csv", index=False)
        df_scenario_options.to_csv(outputFolder + r"\Scenario_options.csv")
        df_conditions.to_csv(outputFolder + r"\conditions_options.csv")

        # Total time of the simulation of this connection
        elapsed = (timeit.default_timer() - start_time) / 60
        print("The Total RunTime is: " + str(elapsed) + " min")


if __name__ == '__main__':
    Main.ask_files()

