###  # -*- coding: iso-8859-15 -*-
# This is my main code.

#Library of Functions for OpenDSS Simulation Results Interpretation

author = "Paulo Radatz and Celso Rocha"
version = "01.00.01"
last_update = "09/27/2017"

import csv
from pylab import *
import operator
import numpy as np
import pandas as pd

from matplotlib.font_manager import FontProperties

fontP = FontProperties()
fontP.set_size('small')



def readVolt_Exceptions(folder):

    df_di_VoltExceptions = pd.read_csv(folder + "/" + r"DI_yr_0\DI_VoltExceptions_1.csv", engine="python")
    # Time - Column 0
    # Undervoltages - Column 1
    # Min Voltage - Column 2
    # Overvoltages - Column 3
    # Max Voltage - Column 4
    # Min Bus - Column 5
    # Max Bus - Column 6
    # LVUndervoltages - Column 7
    # Min LV Voltage - Column 8
    # LVOvervoltages - Column 9
    # Max LV Voltage - Column 10
    # Min LV Bus - Column 11
    # Max LV Bus - Column 12

    sr_minVoltage_value = df_di_VoltExceptions.ix[df_di_VoltExceptions[' "Min Voltage"'].idxmin()]
    sr_maxVoltage_value = df_di_VoltExceptions.ix[df_di_VoltExceptions[' "Max Voltage"'].idxmax()]

    sr_minLVVoltage_value = df_di_VoltExceptions.ix[df_di_VoltExceptions[' "Min LV Voltage"'].idxmin()]
    sr_maxLVVoltage_value = df_di_VoltExceptions.ix[df_di_VoltExceptions[' "Max LV Voltage"'].idxmax()]

    minVoltage_value = sr_minVoltage_value[' "Min Voltage"']
    maxVoltage_value = sr_maxVoltage_value[' "Max Voltage"']
    minLVVoltage_value = sr_minLVVoltage_value[' "Min LV Voltage"']
    maxLVVoltage_value = sr_maxLVVoltage_value[' "Max LV Voltage"']

    underVoltages = sr_minVoltage_value[' "Undervoltages"']
    overVoltages = sr_maxVoltage_value[' "Overvoltage"']
    underLVVoltages = sr_minLVVoltage_value[' "LV Undervoltages"']
    overLVVoltages = sr_maxLVVoltage_value[' "LV Overvoltage"']

    minBus = sr_minVoltage_value[' "Min Bus"']
    maxBus = sr_maxVoltage_value[' "Max Bus"']
    minLVBus = sr_minLVVoltage_value[' "Min LV Bus"']
    maxLVBus = sr_maxLVVoltage_value[' "Max LV Bus"']

    minVoltageTime = sr_minVoltage_value['Hour']
    maxVoltageTime = sr_maxVoltage_value['Hour']
    minLVVoltageTime = sr_minLVVoltage_value['Hour']
    maxLVVoltageTime = sr_maxLVVoltage_value['Hour']


    simulationTimeStep = (float(df_di_VoltExceptions["Hour"].ix[1]) - float(df_di_VoltExceptions["Hour"].ix[0]))

    sr_overvoltage = (df_di_VoltExceptions[' "Overvoltage"'] > 1) | (df_di_VoltExceptions[' "LV Overvoltage"'] > 1)
    sr_undervoltage = (df_di_VoltExceptions[' "Undervoltages"'] > 1) | (df_di_VoltExceptions[' "LV Undervoltages"'] > 1)

    timeAboveANSImaxVoltage = sr_overvoltage[sr_overvoltage == True].count() * simulationTimeStep
    timeBelowANSIminVoltage = sr_undervoltage[sr_undervoltage == True].count() * simulationTimeStep

    resultsList = [maxVoltage_value, overVoltages, maxBus, minVoltage_value, underVoltages, minBus,
                   maxLVVoltage_value, overLVVoltages, maxLVBus, minLVVoltage_value, underLVVoltages, minLVBus,
                   timeAboveANSImaxVoltage, timeBelowANSIminVoltage]

    return resultsList

def read_mon_Power_feederHead(monitor_Power_file):

    # hour - Column 0
    # t(sec) - Column 1
    # P1 - Column 2
    # Q1 - Column 3
    # P2 - Column 4
    # Q2 - Column 5
    # P3 - Column 6
    # Q3 - Column 7
    # P4 - Column 8
    # Q4 - Column 9

    df_mon_file = pd.read_csv(monitor_Power_file, engine="python")


    pt = (df_mon_file[" P1 (kW)"] + df_mon_file[" P2 (kW)"] + df_mon_file[" P3 (kW)"]) * (-1)
    qt = (df_mon_file[" Q1 (kvar)"] + df_mon_file[" Q2 (kvar)"] + df_mon_file[" Q3 (kvar)"]) * (-1)

    # Total PF calculation

    pf = (pt / sqrt(pt ** 2 + qt ** 2))
    pfWOsign = (sqrt(pt ** 2) / sqrt(pt ** 2 + qt ** 2))

    #p1Max = (df_mon_file[" P1 (kW)"] * (-1)).max()
    #p2Max = (df_mon_file[" P2 (kW)"] * (-1)).max()
    #p3Max = (df_mon_file[" P3 (kW)"] * (-1)).max()
    ptMax = ((df_mon_file[" P1 (kW)"] + df_mon_file[" P2 (kW)"] + df_mon_file[" P3 (kW)"]) * (-1)).max()

    #p1Min = (df_mon_file[" P1 (kW)"] * (-1)).min()
    #p2Min = (df_mon_file[" P2 (kW)"] * (-1)).min()
    #p3Min = (df_mon_file[" P3 (kW)"] * (-1)).min()
    ptMin = ((df_mon_file[" P1 (kW)"] + df_mon_file[" P2 (kW)"] + df_mon_file[" P3 (kW)"]) * (-1)).min()

    q1Max = (df_mon_file[" Q1 (kvar)"] * (-1)).max()
    q2Max = (df_mon_file[" Q2 (kvar)"] * (-1)).max()
    q3Max = (df_mon_file[" Q3 (kvar)"] * (-1)).max()
    qtMax = ((df_mon_file[" Q1 (kvar)"] + df_mon_file[" Q2 (kvar)"] + df_mon_file[" Q1 (kvar)"]) * (-1)).max()

    q1Min = (df_mon_file[" Q1 (kvar)"] * (-1)).min()
    q2Min = (df_mon_file[" Q2 (kvar)"] * (-1)).min()
    q3Min = (df_mon_file[" Q3 (kvar)"] * (-1)).min()
    qtMin = ((df_mon_file[" Q1 (kvar)"] + df_mon_file[" Q2 (kvar)"] + df_mon_file[" Q1 (kvar)"]) * (-1)).min()

    pfMin = pf.ix[pfWOsign.idxmin()]

    # Results order
    # 0 - time
    # 1 - p1
    # 2 - q1
    # 3 - p2
    # 4 - q2
    # 5 - p3
    # 6 - q3
    # 7 - pt
    # 8 - qt
    # 9 - pf
    # 10 - pfWOsign
    # 11 - ptMax
    # 12 - ptMin
    # 13 - qtMax
    # 14 - qtMin
    # 15 - pfMin
    # 16 - pfWOsignMin
    # 17 - ptMax_index
    # 18 - ptMin_index

    resultsList = [time, "", "", "", "", "", "", pt, qt, pf, pfWOsign, ptMax, ptMin, qtMax, qtMin, pfMin, "",
                   "", ""]

    return resultsList

def read_mon_Volt(monitor_volt_file, vBase):

    # hour - Column 0
    # t(sec) - Column 1
    # V1 - Column 2
    # VAangle1 - Column 3
    # V2 - Column 4
    # VAngle2 - Column 5
    # V3 - Column 6
    # VAngle3 - Column 7
    # V4 - Column 8
    # VAngle4 - Column 9
    # I1 - Column 10
    # IAangle1 - Column 11
    # I2 - Column 12
    # IAngle2 - Column 13
    # I3 - Column 14
    # IAngle3 - Column 15
    # I4 - Column 16
    # IAngle4 - Column 17

    df_mon_file = pd.read_csv(monitor_volt_file, engine="python")

    time = df_mon_file["hour"] + df_mon_file[" t(sec)"] / 3600.0
    v1 = df_mon_file[" V1"] / vBase
    v2 = df_mon_file[" V2"] / vBase
    v3 = df_mon_file[" V3"] / vBase

    v1Max = v1.max()
    v2Max = v2.max()
    v3Max = v3.max()

    v1Min = v1.min()
    v2Min = v2.min()
    v3Min = v3.min()

    v1Mean = v1.mean()
    v2Mean = v2.mean()
    v3Mean = v3.mean()

    vMax = max(v1Max, v2Max, v3Max)
    vMin = min(v1Min, v2Min, v3Min)
    vMean = np.mean([v1Mean, v2Mean, v3Mean])

    resultsList = [time, v1, v2, v3, vMax, vMin, vMean]

    return resultsList


def read_mons_Volt_to_Calculate_VVI(monitor_volt_base_file, monitor_volt_file):
    # hour - Column 0
    # t(sec) - Column 1
    # V1 - Column 2
    # VAangle1 - Column 3
    # V2 - Column 4
    # VAngle2 - Column 5
    # V3 - Column 6
    # VAngle3 - Column 7
    # V4 - Column 8
    # VAngle4 - Column 9
    # I1 - Column 10
    # IAangle1 - Column 11
    # I2 - Column 12
    # IAngle2 - Column 13
    # I3 - Column 14
    # IAngle3 - Column 15
    # I4 - Column 16
    # IAngle4 - Column 17

    correctCalculation = 1

    df_mon_base_file = pd.read_csv(monitor_volt_base_file, engine="python")
    df_mon_file = pd.read_csv(monitor_volt_file, engine="python")

    vvi_v1 = ((df_mon_file[" V1"].diff() ** 2) ** 0.5).sum() / ((df_mon_base_file[" V1"].diff() ** 2) ** 0.5).sum()
    vvi_v2 = ((df_mon_file[" V2"].diff() ** 2) ** 0.5).sum() / ((df_mon_base_file[" V2"].diff() ** 2) ** 0.5).sum()
    vvi_v3 = ((df_mon_file[" V3"].diff() ** 2) ** 0.5).sum() / ((df_mon_base_file[" V3"].diff() ** 2) ** 0.5).sum()

    vvi = (((df_mon_file[" V1"].diff() ** 2) ** 0.5).sum() + ((df_mon_file[" V2"].diff() ** 2) ** 0.5).sum() + ((df_mon_file[" V3"].diff() ** 2) ** 0.5).sum()) / (((df_mon_base_file[" V1"].diff() ** 2) ** 0.5).sum() + ((df_mon_base_file[" V2"].diff() ** 2) ** 0.5).sum() + ((df_mon_base_file[" V3"].diff() ** 2) ** 0.5).sum())

    if df_mon_file[" V1"].count() == df_mon_base_file[" V1"].count():
        correctCalculation = 0

    resultsList = [vvi_v1, vvi_v2, vvi_v3, vvi, correctCalculation]

    return resultsList

def readTotals(folder):

    # di_Totals = csv.reader(open(self.folder + "/" + r"DI_yr_0\DI_Totals.csv", "r"))
    totals = csv.reader(open(folder + "/" + r"DI_yr_0\Totals_1.csv", "r"))

    # DI_Totals FILE
    # Time - Column 0
    # kWh - Column 1: Circuit Element
    # kvarh - Column 2: Circuit Element
    # Max kW - Column 3: Circuit Element
    # Max kVA - Column 4: Circuit Element
    # Zone kWh - Column 5: Loads
    # Zone kvarh - Column 6: Loads
    # Zone Losses kWh - Column 13
    # Zone Losses kvarh - Column 14

    for row in totals:
        if row[0] != "Year":
            kWhFeederHead = row[1]
            kvarFeederHead = row[2]
            maxkWFeederHead = row[3]
            maxkvarFeederHead = sqrt(float(row[4]) ** 2 - float(row[3]) ** 2)
            kWhFeederConsumption = row[5]
            kvarFeederConsumption = row[6]
            kwhFeederLosses = row[13]
            kvarFeederLosses = row[14]

    resultsList = [kWhFeederHead, kvarFeederHead, maxkWFeederHead, maxkvarFeederHead, kWhFeederConsumption,
                   kvarFeederConsumption, kwhFeederLosses, kvarFeederLosses]

    return resultsList

def operations_counter(monitor_file, start, end):

    # Variables Description:
    #   path: folder path (string)
    #   casename: name of current case being run (string)
    #   monitor: monitor name (string)
    #   start: tuple of 2 elements (hour, sec) for the beginning of the time window to be considered, if set to None, reads from the beginning
    #   end: tuple of 2 elements (hour, sec) for the end of the time window to be considered, if set to None, reads until the end

    csv_current = csv.reader(open(monitor_file, 'r'))
    taps_total = 0
    tap_previous = 0
    for row in csv_current:

        if row[0] != 'hour':
            if start is not None:  # Start reading from the specified time
                if float(row[0]) == start[0] and float(row[1]) == start[1]:
                    tap_previous = float(row[2])

                if float(row[0]) >= start[0] and float(row[1]) > start[1]:

                    tap_current = float(row[2])
                    if tap_current != tap_previous:
                        taps_total = taps_total + 1

                    tap_previous = tap_current

            else:  # Start reading from the beginning of the file
                if csv_current.line_num == 2:
                    tap_previous = float(row[2])
                else:
                    tap_current = float(row[2])
                    if tap_current != tap_previous:
                        taps_total = taps_total + 1

                    tap_previous = tap_current

            if end is not None:   # else, keep reading until the end of the file
                if float(row[0]) == end[0] and float(row[1]) == end[1]:
                    break

    return taps_total

def read_export_Voltage(volt_file):

    df = pd.read_csv(volt_file, engine="python").set_index(keys="Bus", drop=True).drop([" BasekV", " Magnitude1", " Angle1", " Magnitude2", " Angle2", " Magnitude3", " Angle3"], axis=1)

    df_3phase_buses = df[(df[" Node1"] == 1) & (df[" Node2"] == 2) & (df[" Node3"] == 3)]

    df_2phase_buses = df[((df[" Node1"] == 1) & (df[" Node2"] == 2) & (df[" Node3"] == 0)) | ((df[" Node1"] == 2) & (df[" Node2"] == 3) & (df[" Node3"] == 0)) | ((df[" Node1"] == 1) & (df[" Node2"] == 3) & (df[" Node3"] == 0))]

    df_1phase_buses = df[((df[" Node1"] == 1) & (df[" Node2"] == 0) & (df[" Node3"] == 0)) | ((df[" Node1"] == 2) & (df[" Node2"] == 0) & (df[" Node3"] == 0)) | ((df[" Node1"] == 3) & (df[" Node2"] == 0) & (df[" Node3"] == 0))]

    sr_voltage_average = pd.concat([df_3phase_buses[[" pu1", " pu2", " pu3"]].mean(axis=1), df_2phase_buses[[" pu1", " pu2"]].mean(axis=1), df_1phase_buses[[" pu1"]].mean(axis=1)]).astype("float32")

    df_3phase_buses_mean = pd.concat([df_3phase_buses, pd.DataFrame({"Mean": df_3phase_buses[[" pu1", " pu2", " pu3"]].mean(axis=1)})], axis=1)
    df_3phase_buses_mean[" imb1"] = ((((df_3phase_buses_mean[" pu1"] - df_3phase_buses_mean["Mean"])) ** 2) ** 0.5) / df_3phase_buses_mean["Mean"]
    df_3phase_buses_mean[" imb2"] = ((((df_3phase_buses_mean[" pu2"] - df_3phase_buses_mean["Mean"])) ** 2) ** 0.5) / df_3phase_buses_mean["Mean"]
    df_3phase_buses_mean[" imb3"] = ((((df_3phase_buses_mean[" pu3"] - df_3phase_buses_mean["Mean"])) ** 2) ** 0.5) / df_3phase_buses_mean["Mean"]

    sr_voltage_unbalance = df_3phase_buses_mean[[" imb1", " imb2", " imb3"]].max(axis=1).astype("float32")

    return sr_voltage_average, sr_voltage_unbalance

def get_sr_voltages(allNodeNames, allBusVmagPu):

    df_voltages = pd.DataFrame({"Bus": pd.Series(allNodeNames).str.split(".", expand=True)[0], "Node": pd.Series(allNodeNames).str.split(".", expand=True)[1], "V": pd.Series(allBusVmagPu)}).set_index("Bus")
    df_voltages_phases = pd.concat([df_voltages[df_voltages["Node"] == "1"].drop("Node", axis=1), df_voltages[df_voltages["Node"] == "2"].drop("Node", axis=1), df_voltages[df_voltages["Node"] == "3"].drop("Node", axis=1), ], axis=1, sort=False)
    df_voltages_phases.columns = ["v1pu", "v2pu", "v3pu"]
    df_voltages_phases["vMean"] = df_voltages_phases.mean(axis=1)
    df_voltages_phases["v1imb"] = ((((df_voltages_phases["v1pu"] - df_voltages_phases["vMean"])) ** 2) ** 0.5) / df_voltages_phases["vMean"]
    df_voltages_phases["v2imb"] = ((((df_voltages_phases["v2pu"] - df_voltages_phases["vMean"])) ** 2) ** 0.5) / df_voltages_phases["vMean"]
    df_voltages_phases["v3imb"] = ((((df_voltages_phases["v3pu"] - df_voltages_phases["vMean"])) ** 2) ** 0.5) / df_voltages_phases["vMean"]

    return df_voltages_phases["vMean"], df_voltages_phases[["v1imb", "v2imb", "v3imb"]].dropna().max(axis=1)
    #return df_voltages_phases