# ##  # -*- coding: iso-8859-15 -*-
author = "Paulo Radatz"
version = "01.00.02"
last_update = "09/13/2016"

import win32com.client
from win32com.client import makepy
from numpy import *
from pylab import *
import os               # for path manipulation and directory operations
import ReadResults
import pandas as pd
import numpy as np

class DSS(object):

    #------------------------------------------------------------------------------------------------------------------#
    def __init__(self, dssObj, dssFileName, OpenDSS_folder_path, caseResultsPath, connectionID, caseName):

        # OpenDSS Object
        self.dss = dssObj

        self.OpenDSS_folder_path = OpenDSS_folder_path
        self.resultsPath = caseResultsPath
        self.connectionID = connectionID
        self.scenarioName = caseName

        # This variable allows the software print the OpenDSS scripts (best idea is to keep it as an internal variable)
        # User has no access
        self.printOpenDSSscript = "No"

        # Always a good idea to clear the DSS when loading a new circuit
        self.dss.clearAll()
        self.dss.text("clearall")

        # Load the given circuit master file into OpenDSS
        line1 = "compile " + dssFileName
        self.dss.text(line1)

        if self.printOpenDSSscript == "Yes":
            print line1

    def redirect_curves(self, loadinglevel, pvGenCurve):

        line1 = "Redirect " + self.OpenDSS_folder_path + "/LoadShapes_" + loadinglevel + ".dss"
        line2 = "Redirect " + self.OpenDSS_folder_path + "/PVGenCurve_" + pvGenCurve + ".txt"

        self.dss.text(line1)
        if self.printOpenDSSscript == "Yes":
            print line1

        if pvGenCurve is not None:
            self.dss.text(line2)
            if self.printOpenDSSscript == "Yes":
                print line2

    def pvSystems(self, pvLocations, pvSizes_kW):

        vlist = []
        for i in range(len(pvLocations)):
            pv_bus_index = list(self.dss.get_Circuit_AllBusNames()).index(pvLocations[i])
            self.dss.set_Circuit_SetActiveBusi(pv_bus_index)
            kV = self.dss.get_Bus_kVBase() * sqrt(3)

            vlist.append(kV * 1000 / sqrt(3))

            kWPmpp = float(pvSizes_kW[i])
            pvkVA = 1.1 * kWPmpp

            if kWPmpp > 0:
                line1 = "New line.PV_" + str(i) + " phases=3 bus1=" + str(pvLocations[i]) + " bus2=PV_sec_" + str(pvLocations[i]) + " switch=yes"
                line2 = "New transformer.PV_" + str(i) + " phases=3 windings=2 buses=(PV_sec_" + str(pvLocations[i]) + ", PV_ter_" + str(pvLocations[i]) + ") conns=(wye, wye) kVs=(" + str(kV) + ",0.48) xhl=5.67 %R=0.4726 kVAs=(" + str(pvkVA) + "," + str(pvkVA) + ")"          ##kVAs=(4500, 4500)"
                line3 = "makebuslist"
                line4 = "setkVBase bus=PV_sec_" + str(pvLocations[i]) + " kVLL=" + str(kV)
                line5 = "setkVBase bus=PV_ter_" + str(pvLocations[i]) + " kVLL=0.48"
                line6 = "New PVSystem.PV_" + str(pvLocations[i]) + " phases=3 conn=wye  bus1=PV_ter_" + str(pvLocations[i]) + " kV=0.48 kVA=" + str(pvkVA) + " irradiance=1 Pmpp=" + str(kWPmpp) + " pf=1 %cutin=0.05 %cutout=0.05 VarFollowInverter=yes yearly=PVshape kvarlimit=" + str(pvkVA*0.44)
                line7 = "New monitor.terminalV_" + str(i) + " element=transformer.PV_" + str(i) + " terminal=2 mode=0"
                line8 = "New monitor.terminalP_" + str(i) + " element=transformer.PV_" + str(i) + " terminal=2 mode=1 ppolar=no"

                self.dss.text(line1)
                self.dss.text(line2)
                self.dss.text(line3)
                self.dss.text(line4)
                self.dss.text(line5)
                self.dss.text(line6)
                self.dss.text(line7)
                self.dss.text(line8)

                if self.printOpenDSSscript == "Yes":
                    print line1
                    print line2
                    print line3
                    print line4
                    print line5
                    print line6
                    print line7
                    print line8

        return vlist

    def inverter(self, y_curve, x_curve, y_curveW, x_curveW, scenario, pvLocations, pf_dic):

        line1 = "New XYcurve.generic npts=7 yarray=" + y_curve + " xarray=" + x_curve
        line2 = "New XYcurve.genericW npts=4 yarray=" + y_curveW + " xarray=" + x_curveW

        self.dss.text(line1)
        self.dss.text(line2)
        if self.printOpenDSSscript == "Yes":
            print line1
            print line2

        # self.smart_functions = {0: 'PF', 1: 'VV', 2: 'VW', 3: 'DRC', 4: 'VV_VW', 5: 'VV_DRC'}
        for i in range(len(scenario)):

            if scenario[i] == 0:
                line = "PVSystem.PV_" + str(pvLocations[i]) + ".pf=" + pf_dic[i]
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

            elif scenario[i] == 1:
                # Volt/var
                line = "New InvControl.feeder" + str(pvLocations[i]) + " mode=VOLTVAR voltage_curvex_ref=rated vvc_curve1=generic deltaQ_factor=0.2 VV_RefReactivePower=VARMAX_WATTS eventlog=no PVSystemList=PV_" + str(pvLocations[i])
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

            elif scenario[i] == 2:
                # Volt/watt
                line = "New InvControl.feeder" + str(pvLocations[i]) + " mode=VOLTWATT voltage_curvex_ref=rated voltwatt_curve=genericW  VoltwattYAxis=PMPPPU DeltaP_factor=0.45 eventlog=no PVSystemList=PV_" + str(pvLocations[i])
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

            elif scenario[i] == 3:
                # DRC
                line = "New InvControl.feeder" + str(pvLocations[i]) + " Mode=DYNAMICREACCURR DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.15 EventLog=No voltagechangetolerance=0.0001 PVSystemList=PV_" + str(pvLocations[i])
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

            elif scenario[i] == 4:
                # Volt/watt and Volt/var
                line = "New InvControl.feeder" + str(pvLocations[i]) + " CombiMode=VV_VW voltage_curvex_ref=rated voltwatt_curve=genericW  VoltwattYAxis=PMPPPU vvc_curve1=generic deltaQ_factor=0.2 VV_RefReactivePower=VARAVAL_WATT eventlog=no PVSystemList=PV_" + str(pvLocations[i])
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

            elif scenario[i] == 5:
                #Volt/var and DRC
                line = "New InvControl.feeder" + str(pvLocations[i]) + " CombiMode=VV_DRC  voltage_curvex_ref=rated VV_RefReactivePower=varMax_watts vvc_curve1=generic DbVMin=1 DbVMax=1 ArGraLowV=50  arGraHiV=50 DynReacavgwindowlen=300s  deltaQ_factor=.15 EventLog=No voltagechangetolerance=0.0001 PVSystemList=PV_" + str(pvLocations[i])
                self.dss.text(line)
                if self.printOpenDSSscript == "Yes":
                    print line

 #------------------------------------------------------------------------------------------------------------------#

    #------------------------------------------------------------------------------------------------------------------#
    def solve_yearly(self):

        # Simulation Options
        stepsize = 5
        number_get_voltage = 180
        tSinterval = 17280  # 5759;%4320; 17280 8640

        get_voltages = self.get_voltages_list(number_get_voltage, tSinterval)

        self.sr_voltage_average_dic = {}
        self.sr_voltage_unbalance_dic = {}
        volt_file = self.resultsPath + r"/" + self.scenarioName + "_EXP_VOLTAGES.csv"

        line1 = "set DataPath=" + self.resultsPath #+ self.scenarioName
        line2 = "set Casename=" + self.scenarioName

        self.dss.text(line1)
        self.dss.text(line2)

        if self.printOpenDSSscript == "Yes":
            print line1
            print line2

        line1 = 0
        line2 = "batchedit load..* yearly=Residential" ###olhar esse ponto e comentar
        line3 = "set maxcontroliter=2000"
        line4 = "set maxiterations=100"
        line5 = "Solve mode=yearly controlmode=static stepsize=" + str(stepsize) + "s number=1"
        line66 = "Reset Monitors"
        line77 = "Reset Meters"
        line6 = "set DemandInterval=true"
        line7 = "set overloadreport=true"
        line8 = "set voltexceptionreport=true"
        line9 = "set DIVerbose=true"
        line10 = "set controlmode=time"
        #line11 = "set number=" + str(tSinterval)
        #line12 = "solve"
        #line11 = "set number=" + str(number_get_voltage)

        self.dss.set_AllowForms(line1)
        self.dss.text(line2)
        self.dss.text(line3)
        self.dss.text(line4)
        self.dss.text(line5)
        #self.dss.text(line66)
        self.dss.set_Meters_ResetAll()
        self.dss.set_Monitors_ResetAll()
        #self.dss.text(line77)
        self.dss.text(line6)
        self.dss.text(line7)
        self.dss.text(line8)
        self.dss.text(line9)
        self.dss.text(line10)
        #self.dss.text(line11)
        #solved = self.dss.set_Solution_Solve()

        if self.dss.get_Meters_DIFilesAreopen() == 0:
            self.dss.set_Meters_OpenAllDIFiles()

        for stepSimulation in range(tSinterval-1):

            control_iter = 0  # FUNCTION TSolutionObj.SolveSnap:Integer;
            self.dss.set_Solution_InitSnap()  # Inicializa Contadores

            while self.dss.get_Solution_ControlActionsDone() == 0:
                solved = self.dss.set_Solution_SolveNoControl()
                solved = self.dss.get_Solution_SampleControlDevices()
                solved = self.dss.set_Solution_DoControlActions()

                # solved = self.dss.set_Solution_SolvePlusControl()

                control_iter += 1
                if control_iter >= self.dss.get_Solution_MaxControlIterations():
                    print u"Número máximo de iterações de controle excedido!"
                    print u"No instante hora:" + str(self.dss.get_Solution_Hour()) ####+ u' e segundo: ' + str(self.dssSolution.Seconds)
                    # self.dssText.Command = "show eventlog"
                    exit()
                    break

            self.dss.get_Meters_SampleAll()
            self.dss.get_Solution_FinishTimeStep()

            if (stepSimulation + 1) in get_voltages:

                sr_voltage_average, sr_voltage_unbalance = self.get_sr_voltages_dll(self.dss.get_Circuit_AllNodeNames(), self.dss.get_Circuit_AllBusVMagPu())
                self.sr_voltage_average_dic[stepSimulation] = sr_voltage_average
                self.sr_voltage_unbalance_dic[stepSimulation] = sr_voltage_unbalance

        self.dss.set_Meters_CloseAllDIFiles()

        if self.printOpenDSSscript == "Yes":
            #print line1
            print line2
            print line3
            print line4
            print line5
            print line6
            print line7
            print line8
            print line9
            print line10
            #print line11
            #print line12
            #print line13
            #print line14

    def get_voltages_list(self, number_get_voltage, tSinterval):

        i = 0
        get_voltages = []
        while (i * number_get_voltage) < tSinterval:
            get_voltages.append(i * number_get_voltage)

            i = i + 1
        return get_voltages

    def get_sr_voltages_dll(self, allNodeNames, allBusVmagPu):

        sr_voltage_average, sr_voltage_unbalance = ReadResults.get_sr_voltages(allNodeNames, allBusVmagPu)

        return sr_voltage_average, sr_voltage_unbalance

    def get_sr_voltages_file(self, volt_file):

        sr_voltage_average, sr_voltage_unbalance = ReadResults.read_export_Voltage(volt_file)

        return sr_voltage_average, sr_voltage_unbalance


    def export_Monitors(self, pvLocations, feederHead_monVolt, feederHead_monPower, regulators_monTaps, capControls_monCaps, feederEnd_monVolt):

        num = len(pvLocations)
        for i in range(num):
            line1 = "export monitors terminalV_" + str(i)
            line2 = "export monitors terminalP_" + str(i)
            self.dss.text(line1)
            self.dss.text(line2)

            if self.printOpenDSSscript == "Yes":
                print line1
                print line2

        for i in range(len(feederHead_monVolt)):
            line1 = "export monitors FeederHeadV"
            line2 = "export monitors FeederHeadP"

            self.dss.text(line1)
            self.dss.text(line2)

            if self.printOpenDSSscript == "Yes":
                print line1
                print line2

        if regulators_monTaps is not None:
            for i in range(len(regulators_monTaps)):
                line1 = "export monitors Regulator_" + str(i)

                self.dss.text(line1)

                if self.printOpenDSSscript == "Yes":
                    print line1

        if capControls_monCaps is not None:
            for i in range(len(capControls_monCaps)):
                line1 = "export monitors SwitchedCap_" + str(i)

                self.dss.text(line1)

                if self.printOpenDSSscript == "Yes":
                    print line1

        for i in range(len(feederEnd_monVolt)):
            line1 = "export monitors FeederEndV"

            self.dss.text(line1)

            if self.printOpenDSSscript == "Yes":
                print line1

    def compileMonitors(self, feederHead_monVolt, feederHead_monPower, regulators_monTaps, capControls_monCaps, feederEnd_monVolt):

        for i in range(len(feederHead_monVolt)):
            line1 = feederHead_monVolt[i]
            line2 = feederHead_monPower[i]

            self.dss.text(line1)
            self.dss.text(line2)

            if self.printOpenDSSscript == "Yes":
                print line1
                print line2

        if regulators_monTaps is not None:
            for i in range(len(regulators_monTaps)):
                line1 = regulators_monTaps[i]

                self.dss.text(line1)

                if self.printOpenDSSscript == "Yes":
                    print line1

        if capControls_monCaps is not None:
            for i in range(len(capControls_monCaps)):
                line1 = capControls_monCaps[i]

                self.dss.text(line1)

                if self.printOpenDSSscript == "Yes":
                    print line1

        for i in range(len(feederEnd_monVolt)):
            line1 = feederEnd_monVolt[i]

            self.dss.text(line1)

            if self.printOpenDSSscript == "Yes":
                print line1

    def setFeederHeadMonitors(self):

        self.dss.set_Vsources_First()
        basev = self.dss.get_Vsources_BasekV() * 1000 / sqrt(3)   # Tensao de Base de Fase em Volts
        mon_pot = ["New monitor.FeederHeadP element=Vsource.source terminal=1 mode=1 ppolar=no"]     # as potencias aparecem com o sinal invertido.
        mon_volt = ["New monitor.FeederHeadV element=Vsource.source terminal=1 mode=0"]

        return mon_pot, mon_volt, basev

    def setFeederEndMonitors(self):

        medium_v_buses_names = []
        medium_v_buses_distances = []

        for i in range(len(list(self.dss.get_Circuit_AllBusNames()))):
            self.dss.set_Circuit_SetActiveBusi(i)
            if self.dss.get_Bus_kVBase() >= 1 and self.dss.get_Bus_NumNodes() == 3:
                medium_v_buses_names.append(self.dss.get_Bus_Name())
                medium_v_buses_distances.append(self.dss.get_Bus_Distance())

        end_bus_name = medium_v_buses_names[medium_v_buses_distances.index(max(medium_v_buses_distances))]
        end_bus_index = list(self.dss.get_Circuit_AllBusNames()).index(end_bus_name)
        self.dss.set_Circuit_SetActiveBusi(end_bus_index)
        end_bus_vbase = self.dss.get_Bus_kVBase() * 1000    # Tensao de Base de Fase em Volts

        # Agora, e preciso achar a linha a qual essa barra esta conectada para poder declarar o monitor
        self.dss.set_Lines_First()
        for i in range(self.dss.get_Lines_Count()):
            if self.dss.get_Lines_Bus1().split('.')[0] == end_bus_name:
                mon_volt = ["New monitor.FeederEndV element=Line." + self.dss.get_Lines_Name() + " terminal=1 mode=0"]
                break
            elif self.dss.get_Lines_Bus2().split('.')[0] == end_bus_name:
                mon_volt = ["New monitor.FeederEndV element=Line." + self.dss.get_Lines_Name() + " terminal=2 mode=0"]
                break

            self.dss.set_Lines_Next()

        #mon_volt = [u'New monitor.FeederEndV element=Line.0971229_0971230 terminal=2 mode=0']
        #end_bus_vbase = 6928.20323028

        return mon_volt, end_bus_vbase

    def setRegulatorsMonitors(self):
        mon_reg = []
        self.dss.set_RegControls_First()
        for i in range(self.dss.get_RegControls_Count()):
            mon_reg.append("New monitor.Regulator_" + str(i) + " element=transformer." + self.dss.get_RegControls_Transformer() + " terminal=" + str(self.dss.get_RegControls_TapWinding()) + " mode=2")
            self.dss.set_RegControls_Next()
        if mon_reg:
            return mon_reg
        else:
            return None


    def setCapacitorsMonitors(self):
        mon_cap_controls = []
        self.dss.set_CapControls_First()
        for i in range(self.dss.get_CapControls_Count()):
            mon_cap_controls.append("New monitor.SwitchedCap_" + str(i) + " element=capacitor." + self.dss.get_CapControls_Capacitor() + " terminal=1 mode=6")
            self.dss.set_CapControls_Next()
        if mon_cap_controls:
            return mon_cap_controls
        else:
            return None