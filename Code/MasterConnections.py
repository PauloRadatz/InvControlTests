author = "Paulo Radatz and Celso Rocha"
version = "01.00.03"
last_update = "10/13/2017"

"git information: https://confluence.jetbrains.com/display/PYH/Using+PyCharm%27s+Git+integration+locally"

import MasterSettings
import COMinterfaceConnected
import os
import csv
import timeit
import Tkinter as tk
import tkFileDialog
import tkMessageBox
import datetime
import pandas as pd
import sqlite3
import DSS
from sqlalchemy.types import String

class Connections(object):

    ConnectionList = []

    @classmethod
    def ask_files(cls):

        # Way to close the window that appears with no reason
        root = tk.Tk()
        root.withdraw()

        #print "\nPlease select a OpenDSS Feeder Model."
        #dssFileName = tkFileDialog.askopenfilename()
        dssFileName = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/Creelman971_VVDRC/Master_NoPV.dss'
        #dssFileName = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/Green/Master_NoPV.dss'
        #dssFileName = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/ValleyCenter909/Master_NoPV.dss'
        #print "\nPlease select a File with the Connections Selected (*.csv file)."
        #scenarioConfiguration_file = tkFileDialog.askopenfilename()
        scenarioConfiguration_file = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/Creelman971_VVDRC/connections-TimeSeries_HC.csv'
        #scenarioConfiguration_file = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/Green/connections-TimeSeries_HC.csv'
        #scenarioConfiguration_file = u'D:/Projetos_GitHub/Combining_Smart_Inverter_Functions/Feeders/ValleyCenter909/connections-TimeSeries_HC.csv'

        sqlite_file = u"D:\Projetos_GitHub\Combining_Smart_Inverter_Functions\Results\Results.db"
        try:
            config = cls(scenarioConfiguration_file, dssFileName, sqlite_file)
            print "Finished" + str(config)
            tkMessageBox.showinfo("", "Finished")
        except Exception as e:
            raise
            print "Incomplete"
            tkMessageBox.showwarning("", "Incomplete")
        return

    def __init__(self, fileName, dssFileName, sqlite_file):
        """
        :param fileName: This is a .csv file with the connections selected
        :param dssFileName: This is a Master.dss file. The circuit element should be defined in this file
        :param sqlite_file: This a database file. The connection results is stored in there
        """

        # Dictionary for connection DataFrame: Key=Connection Id and Values=DataFrame
        df_connection_results_dic = {}

        # Connect Python to OpenDSS via dll. It is done just once in the code
        self.dssObject = DSS.DSS()

        # OpenDSS Model Directory
        self.OpenDSS_folder_path = os.path.dirname(dssFileName)

        # Data and time of the start of simulation
        startSimulation = str(datetime.datetime.now())

        # Start timing of the whole simulation
        start_time = timeit.default_timer()

        print "Simulation starts: " + str(startSimulation)

        # Read the connections file
        df_connections = pd.read_csv(fileName, engine="python")

        # Creates Connection Objects
        for index, row in df_connections.iterrows():
            print u"Connection ID " + str(row[1]) + u" Created"
            # Each connection is an object of the Settings class
            connectionObject = MasterSettings.Settings(self.dssObject, dssFileName, row)
            Connections.ConnectionList.append(connectionObject)



        # Runs Connection Objects
        for connectionObject in Connections.ConnectionList:
            print u"\n-----------------Running Connection ID " + str(connectionObject.connectionID) + u"--------------------"
            print u"Penetration Level (kW) = " + str(connectionObject.penetrationLevel)
            print u"Number of PV sites = " + str(connectionObject.numberPV)
            print u"Solar profile = " + connectionObject.pvGenCurve
            print u"Loading level = " + connectionObject.loadinglevel

            # Start timing of the connection ID simulation
            start_time_connection = timeit.default_timer()
            connectionObject.process()
            print u"\nSimulation time Connection ID " + str(connectionObject.connectionID) + " = " + str((timeit.default_timer() - start_time_connection)/60) + u" min"

            # Makes a database connection
            conn = sqlite3.connect(sqlite_file)
            print u"\n-----------------Reading Connection ID " + str(connectionObject.connectionID) + u" Results--------------------"
            df_connection_results_dic[str(connectionObject.connectionID)] = pd.DataFrame(connectionObject.scenarioResultsList)

            # https://andypicke.github.io/sql_pandas_post/    , chunksize=1000, dtype={"Col_21": String}
            df_connection_results_dic[str(connectionObject.connectionID)].to_sql(connectionObject.feederName + " C" + str(connectionObject.connectionID), conn, if_exists="replace", index_label="Scenario ID")

            conn.close()

        # Total time of the simulation of this connection
        elapsed = (timeit.default_timer() - start_time) / 60 / 60
        print "The Total RunTime is: " + str(elapsed) + " Hours"


if __name__ == '__main__':
    Connections.ask_files()

