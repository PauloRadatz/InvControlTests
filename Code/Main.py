# encoding: utf-8

import Scenario
import DSS
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

        print "\nPlease select a File with the Connections Selected (*.csv file)."
        #conditionConfiguration_file = tkFileDialog.askopenfilename()
        conditionConfiguration_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\input.csv"


        sqlite_file = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\13Bus\Results.db"
        try:
            config = cls(conditionConfiguration_file, dssFileName, sqlite_file)
            print("Finished" + str(config))
            tkMessageBox.showinfo("", "Finished")
        except Exception as e:
            raise
            print("Incomplete")
            tkMessageBox.showwarning("", "Incomplete")
        return

    def __init__(self, fileName, dssFileName, sqlite_file):

        df_conditions_results_dic = pd.DataFrame()


        # Connect Python to OpenDSS via dll or COM. It is done just once in the code
        self.dssObject = DSS.DSS()

        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

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
            conditionObject = Scenario.Settings(self.dssObject, dssFileName, row)
            Main.ConditionList.append(conditionObject)

        # Runs Connection Objects
        for condition in Main.ConditionList:
            print(u"\n-----------------Running Condition ID " + str(condition.conditionID) + u"--------------------")
            print(u"Penetration Level (%) = " + str(condition.penetrationLevel))
            #print(u"Number of PV sites = " + str(condition.numberPV))

            # Start timing of the connection ID simulation
            start_time_connection = timeit.default_timer()
            for i in range(condition.numberScenarios):
                condition.process()

            print(u"\nSimulation time Condition ID " + str(condition.conditionID) + " = " + str((timeit.default_timer() - start_time_connection)/60) + u" min")

            # Makes a database Condition
            conn = sqlite3.connect(sqlite_file)
            print(u"\n-----------------Reading Condition ID " + str(condition.conditionID) + u" Results--------------------")
            ##df_conditions_results_dic[str(condition.conditionID)] = pd.DataFrame(condition.scenarioResultsList)

            # https://andypicke.github.io/sql_pandas_post/    , chunksize=1000, dtype={"Col_21": String}
            ##df_conditions_results_dic[str(condition.conditionID)].to_sql(condition.feederName + " C" + str(condition.conditionID), conn, if_exists="replace", index_label="Scenario ID")

            conn.close()

        # Total time of the simulation of this connection
        elapsed = (timeit.default_timer() - start_time) / 60 / 60
        print("The Total RunTime is: " + str(elapsed) + " Hours")


if __name__ == '__main__':
    Main.ask_files()

