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
    list_false = ["0", "No", "no", "n", "NO", 0, 0.0]
    list_default = ["Default", "default", "d"]
    list_true = ["1", "Yes", "yes", "ny", "YES", 1, 1.0]

    # @classmethod
    # def ask_files(cls):
    #
    #     # Way to close the window that appears with no reason
    #     root = tk.Tk()
    #     root.withdraw()
    #
    #     print "\nPlease select a OpenDSS Feeder Model."
    #     #dssFileName = tkFileDialog.askopenfilename()
    #     dssFileName = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\IEEE13Nodeckt.dss"
    #     #dssFileName = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Creelman\Master_NoPV.dss"
    #     #dssFileName = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\123Bus\IEEE123Master.dss"
    #
    #     print "\nPlease select the definitions of the Conditions (*.csv file)."
    #     #conditionConfiguration_file = tkFileDialog.askopenfilename()
    #     conditionConfiguration_file = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input.csv"
    #     #conditionConfiguration_file = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input_13.csv"
    #     #conditionConfiguration_file = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Creelman\Results\Delta_PFactor.csv"
    #     #conditionConfiguration_file = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\123Bus\input_123.csv"
    #
    #     print "\nPlease select the Output Folder."
    #     #outputFolder = tkFileDialog.askdirectory()
    #     #outputFolder = r"D:\test-1"
    #     #outputFolder = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Creelman\Results\Delta_PFactor"
    #
    #     try:
    #         config = cls(conditionConfiguration_file, dssFileName, outputFolder)
    #         print("Finished" + str(config))
    #         tkMessageBox.showinfo("", "Finished")
    #     except Exception as e:
    #         raise
    #         print("Incomplete")
    #         tkMessageBox.showwarning("", "Incomplete")
    #     return

    def __init__(self, fileName, dssFileName, outputFolder, outputFolder_temp, scenario):

        ConditionList = []

        # Data and time of the start of simulation
        startSimulation = str(datetime.datetime.now())

        # Start timing of the whole simulation
        start_time = timeit.default_timer()

        print("Simulation starts: " + str(startSimulation))

        # Scenario Options file
        #df_scenario_options = pd.read_csv(os.path.dirname(fileName) + r"\Scenario_options.csv", engine="python").set_index("Option")
        df_scenario_options = pd.read_csv("D:\Projetos_GitHub\InvControlTests\Scenario_options_" + scenario + ".csv", engine="python").set_index("Option")

        # Read the conditions file
        df_conditions = pd.read_csv(fileName, engine="python")

        # Creates a methodology object
        methodologyObj = Methodology.Methodology(outputFolder_temp)

        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        # DataFrame that will store scenarios of the first (base) condition
        df_scenarios_fixed = ""
        base = "yes"

        if not os.path.exists(outputFolder_temp):
            os.makedirs(outputFolder_temp)

        # Creates condition Objects
        for index, row in df_conditions.iterrows():
            print(u"Scenario ID " + str(row[1]) + u" Created")
            # Each connection is an object of the Settings class
            conditionObject = Scenario.Settings(dssFileName, row, methodologyObj, df_scenario_options)
            ConditionList.append(conditionObject)

        # Runs Connection Objects
        for condition in ConditionList:
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
    # Main.ask_files()

    dssFileName = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Creelman\Master_NoPV.dss"
    #dssFileName = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\123Bus\IEEE123Master.dss"



    #input = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Input\Creelman"
    input = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Input_QSTS\Creelman"
    #input = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Input\123Bus"
    #input = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Input_QSTS\123Bus"
    cases_list = [r"\1-50_10_Q",    r"\2-50_10_P",    r"\3-50_10_All",
                  r"\4-50_50_Q",    r"\5-50_50_P",    r"\6-50_50_All",
                  r"\7-50_100_Q",   r"\8-50_100_P",   r"\9-50_100_All",
                  r"\10-100_10_Q",  r"\11-100_10_P",  r"\12-100_10_All",
                  r"\13-100_50_Q",  r"\14-100_50_P",  r"\15-100_50_All",
                  r"\16-100_100_Q", r"\17-100_100_P", r"\18-100_100_All",
                  r"\19-150_10_Q",  r"\20-150_10_P",  r"\21-150_10_All",
                  r"\22-150_50_Q",  r"\23-150_50_P",  r"\24-150_50_All",
                  r"\25-150_100_Q", r"\26-150_100_P", r"\27-150_100_All"]

    #output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_01"
    #output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\123Bus_deltas"

    output_temp = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_temp_QSTS\Creelman"
    #output_temp = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_temp_QSTS\123Bus"

    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_NotUpdate"
    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_Update"
    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_UpdateP"
    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_02"
    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_01"
    output = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_005"

    for case in cases_list:
        fileName = input + case + ".csv"
        outputfolder = output + case
        outputfolder_temp = output_temp + case
        scenario = case.split("_")[-1]
        Main(fileName, dssFileName, outputfolder, outputfolder_temp, scenario)


