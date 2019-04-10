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

    @classmethod
    def ask_files(cls):

        # Way to close the window that appears with no reason
        root = tk.Tk()
        root.withdraw()

        print "\nPlease select a OpenDSS Feeder Model."
        #dssFileName = tkFileDialog.askopenfilename()
        dssFileName = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\IEEE13Nodeckt.dss"

        print "\nPlease select the Conditions (*.csv file)."
        #conditionConfiguration_file = tkFileDialog.askopenfilename()
        conditionConfiguration_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input.csv"

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

        # sqlite_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\Results.db"

        df_conditions_results_dic = pd.DataFrame()

        methodologyObj = Methodology.Methodology()

        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        # Options
        df_options = pd.read_csv(os.path.dirname(fileName) + r"\Options.csv", engine="python").set_index("Option")
        # Scenario Options file
        df_scenario_options = pd.read_csv(os.path.dirname(fileName) + r"\Scenario_options.csv", engine="python").set_index("Option")

        # Data and time of the start of simulation
        startSimulation = str(datetime.datetime.now())

        # Start timing of the whole simulation
        start_time = timeit.default_timer()

        print("Simulation starts: " + str(startSimulation))

        # Read the conditions file
        df_conditions = pd.read_csv(fileName, engine="python")

        # Creates condition Objects
        for index, row in df_conditions.iterrows():
            print(u"Scenario ID " + str(row[1]) + u" Created")
            # Each connection is an object of the Settings class
            conditionObject = Scenario.Settings(dssFileName, row, methodologyObj, df_scenario_options)
            Main.ConditionList.append(conditionObject)

        df_buses_fixed = ""
        base = "yes"

        # Runs Connection Objects
        for condition in Main.ConditionList:
            print(u"\n-----------------Running Condition ID " + str(condition.conditionID) + u"--------------------")
            print(u"Penetration Level (%) = " + str(condition.penetrationLevel))

            condition_outputFolder = outputFolder + "/" + str(condition.conditionID)

            if not os.path.exists(condition_outputFolder):
                os.makedirs(condition_outputFolder)

            sqlite_file = condition_outputFolder + r"\Scenarios_issue.db"

            # Start timing of the connection ID simulation
            start_time_connection = timeit.default_timer()

            for i in range(condition.numberScenarios):
                condition.process(i, df_buses_fixed, base, df_options["Value"]["Scenarios Buses"])
            if base == "yes":
                df_buses_fixed = condition.df_buses_scenarios
                base = "no"

            condition.get_results()

            print(u"\nSimulation time Condition ID " + str(condition.conditionID) + " = " + str((timeit.default_timer() - start_time_connection)/60) + u" min")

            # Makes a database Condition
            conn = sqlite3.connect(sqlite_file)
            print(u"\n-----------------Reading Condition ID " + str(condition.conditionID) + u" Results--------------------")

            for issue in condition.df_issue_dic:
                condition.df_issue_dic[issue].to_sql(str(issue), conn, if_exists="replace")
                condition.df_issue_dic[issue].to_csv(condition_outputFolder + "/" + str(issue) + ".csv")


            ##df_conditions_results_dic[str(condition.conditionID)] = pd.DataFrame(condition.scenarioResultsList)

            # https://andypicke.github.io/sql_pandas_post/    , chunksize=1000, dtype={"Col_21": String}
            ##df_conditions_results_dic[str(condition.conditionID)].to_sql(condition.feederName + " C" + str(condition.conditionID), conn, if_exists="replace", index_label="Scenario ID")

            conn.close()

        df_condition_results = pd.DataFrame().from_dict(methodologyObj.resultsObj.sr_condition_statistics_results_dic)

        df_condition_results.T.to_csv(outputFolder + r"\condition_results.csv", index=False)
        df_scenario_options.to_csv(outputFolder + r"\Scenario_options.csv")
        df_conditions.to_csv(outputFolder + r"\conditions_options.csv")



        # Total time of the simulation of this connection
        elapsed = (timeit.default_timer() - start_time) / 60 / 60
        print("The Total RunTime is: " + str(elapsed) + " Hours")


if __name__ == '__main__':
    Main.ask_files()

