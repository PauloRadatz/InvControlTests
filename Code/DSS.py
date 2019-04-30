# -*- coding: iso-8859-15 -*-


import ctypes
import win32com.client
from win32com.client import makepy
# from numpy import *
from pylab import *
from comtypes import automation
import os
import sys

class DSS(object):

    def __init__(self, tipo=None):
        """ Cria um novo objeto DSS """

        if tipo == 1:
            # These variables provide direct interface into OpenDSS
            # sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
            # makepy.main()  # ensures early binding and improves speed

            # Create a new instance of the DSS
            self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

            # Start the DSS
            if not self.dssObj.Start(0):
                ctypes.windll.user32.MessageBoxW(0, u"Nao foi possível realizar a conexao com o OpenDSS",
                                                 u"Conexao OpenDSS", 0)

            # Acesso direto para as interfaces do OpenDSS
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssMeters = self.dssCircuit.Meters
            self.dssMonitors = self.dssCircuit.Monitors
            self.dssPDElement = self.dssCircuit.PDElements
            self.dssSource = self.dssCircuit.Vsources
            self.dssTransformer = self.dssCircuit.Transformers
            self.dssLines = self.dssCircuit.Lines
            self.dssLoads = self.dssCircuit.Loads
            self.dssRegulators = self.dssCircuit.RegControls
            self.dssCapControls = self.dssCircuit.CapControls
            self.dssTopology = self.dssCircuit.Topology

        else:

            os.chdir(os.path.dirname(sys.argv[0]))
            self.dssObj = ctypes.WinDLL("OpenDSSDirect.dll")
            self.AllocateMemory()

            if int(self.dssObj.DSSI(ctypes.c_int32(3), ctypes.c_int32(0))) == 1:
                versao = ctypes.c_char_p(self.dssObj.DSSS(ctypes.c_int32(1), "".encode('ascii')))
                print u"Conexão com o OpenDSSDirect.dll realizada com sucesso! Versão: " + versao.value.decode('ascii')

            # Desativa formulários
            self.dssObj.DSSI(ctypes.c_int32(8), ctypes.c_int32(0))

    def AllocateMemory(self):
        """ Método que arruma o problema de alocação de memória para interfaces do tipo Float """
        self.dssObj.TransformersF.restype = ctypes.c_double
        self.dssObj.BUSF.restype = ctypes.c_double
        self.dssObj.DSSLoadsF.restype = ctypes.c_double
        self.dssObj.DSSLoadsS.restype = ctypes.c_char_p
        self.dssObj.DSSS.restype = ctypes.c_char_p
        # self.dssObj.DSSPut_Command = ctypes.c_char_p
        self.dssObj.CircuitS.restype = ctypes.c_char_p
        self.dssObj.TopologyS.restype = ctypes.c_char_p
        self.dssObj.MetersS.restype = ctypes.c_char_p
        self.dssObj.RegControlsS.restype = ctypes.c_char_p
        self.dssObj.CapControlsS.restype = ctypes.c_char_p
        self.dssObj.DSSLoadsS.restype = ctypes.c_char_p
        self.dssObj.BUSS.restype = ctypes.c_char_p
        self.dssObj.TransformersS.restype = ctypes.c_char_p
        self.dssObj.DSSLoadsF.restype = ctypes.c_double
        self.dssObj.VsourcesF.restype = ctypes.c_double
        self.dssObj.CircuitS.restype = ctypes.c_char_p
        # self.dssObj.SolutionI.restype = ctypes.c_uint64


# DSS Interface

    # DSSI (int)
    def clearAll(self):
        """ Método que limpa a memória do OpenDSSDirect.dll"""
        self.dssObj.DSSI(ctypes.c_int32(1), ctypes.c_int32(0))

    def set_AllowForms(self, parameter):
        """ Método que desabilita as telas de erros"""
        self.dssObj.DSSI(ctypes.c_int32(8), ctypes.c_int32(parameter))

    # DSSS (String)
    # DSSV (Variant)

# CktElement Interface

    # CktElementI (int)
    def get_CktElement_NumPhases(self):
        """ Método que retorna a quantidade de fases do elemento ativo"""
        resposta = int(self.dssObj.CktElementI(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta

    # CktElementF (Float)
    # CktElementS (String)
    # CktElementV (Variant)
    def get_CktElement_BusNames(self):
        """ Método que retorna uma lista com o nome das barras à qual o elemento ativo se conecta"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CktElementV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

# Circuit Interface

    # CircuitI (int)
    def set_Circuit_SetActiveBusi(self, i):
        """ Método que ativa a barra i"""
        resposta = self.dssObj.CircuitI(ctypes.c_int32(9), ctypes.c_int32(i))
        return resposta

    # CircuitF (Float)
    # CircuitS (String)
    def set_Circuit_ActiveElement_Name(self, nome):
        """ Método que torna um elemento ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(3), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def set_Circuit_ActiveBus_Name(self, nome):
        """ Método que torna uma barra ativa pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(4), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def set_Circuit_SetActiveElement(self, nome):
        """ Método que ativa a barra i"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(0), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    # CircuitV (Variant)
    def get_Circuit_AllBusNames(self):
        """ Método que get os nomes das barras"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(7), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def get_Circuit_AllNodeNames(self):
        """ Método que get os nomes dos nós"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(10), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def get_Circuit_AllBusVMagPu(self):
        """ Método que retorna as tensões das barras em pu"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(9), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def get_Circuit_TotalPower(self):
        """ Método que retorna as potências do circuito"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(3), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

# Bus Interface

    # BusI (int)
    def get_Bus_NumNodes(self):
        """ Método que retorna a quantidade de nós da barra ativa"""
        resposta = int(self.dssObj.BUSI(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta

    # BusF (Float)
    def get_Bus_kVBase(self):
        """ Método que retorna o kVBase da barra ativa"""
        resposta = float(self.dssObj.BUSF(ctypes.c_int32(0), ctypes.c_double(0)))
        return resposta

    def get_Bus_Distance(self):
        """ Método que retorna a distancia do energymeter da barra ativa"""
        resposta = float(self.dssObj.BUSF(ctypes.c_int32(5), ctypes.c_double(0)))
        return resposta

    # BusS (String)
    def get_Bus_Name(self):
        """ Método que retorna o nome da barra ativa"""
        resposta = ctypes.c_char_p(self.dssObj.BUSS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # BusV (Variant)
    def get_Bus_Nodes(self):
        """ Método que retorna os nós da barra ativa"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.BUSV(ctypes.c_int32(2), variant_pointer)
        return variant_pointer.contents.value

# Solution Interface

    # SolutionI (int)
    def set_Solution_Solve(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta

    def get_Solution_Hour(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(3), ctypes.c_int32(0)))
        return resposta

    def get_Solution_Converged(self):
        """ Método que indica se a solução do circuito convergiu ou não. Retorna 1, caso positivo e 0, caso negativo"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(38), ctypes.c_int32(0)))
        return resposta

    def set_Solution_InitSnap(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(32), ctypes.c_int32(0)))
        return resposta

    def get_Solution_ControlActionsDone(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(42), ctypes.c_int32(0)))
        return resposta

    def set_Solution_SolveNoControl(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(30), ctypes.c_int32(0)))
        return resposta

    def set_Solution_SolvePlusControl(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(31), ctypes.c_int32(0)))
        return resposta

    def get_Solution_SampleControlDevices(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(34), ctypes.c_int32(0)))
        return resposta

    def set_Solution_DoControlActions(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(35), ctypes.c_int32(0)))
        return resposta

    def get_Solution_MaxControlIterations(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(24), ctypes.c_int32(0)))
        return resposta

    def get_Solution_FinishTimeStep(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(44), ctypes.c_int32(0)))
        return resposta

    def get_Solution_ControlIterations(self):
        """ PR"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(22), ctypes.c_int32(0)))
        return resposta

    # SolutionF (Float)
    # SolutionS (String)
    # SolutionV (Variant)

# Monitor Interface

    # MonitorsI (int)
    def set_Monitors_ResetAll(self):
        """ Método que reseta todos os monitores do circuito"""
        resposta = int(self.dssObj.MonitorsI(ctypes.c_int32(3), ctypes.c_int32(0)))
        return resposta

# Meters Interface

    # MetersI (int)
    def set_Meters_ResetAll(self):
        """ Método que limpa todos os medidores"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(3), ctypes.c_int32(0))
        return resposta

    def get_Meters_DIFilesAreopen(self):
        """ PR"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(8), ctypes.c_int32(0))
        return resposta

    def set_Meters_OpenAllDIFiles(self):
        """ PR"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(11), ctypes.c_int32(0))
        return resposta

    def set_Meters_CloseAllDIFiles(self):
        """ PR"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(12), ctypes.c_int32(0))
        return resposta

    def get_Meters_SampleAll(self):
        """ PR"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(9), ctypes.c_int32(0))
        return resposta

    # MetersS (String)
    def set_Meters_ActiveName(self, nome):
        """ Método que seta um medidor ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.MetersS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def get_Meters_ActiveName(self):
        """ Método que retorna o nome do medidor ativo"""
        resposta = ctypes.c_char_p(self.dssObj.MetersS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # MetersV (Variant)
    def get_Meters_AllNames(self):
        """ Método que retorna uma lista com o nome de todos os medidores do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def get_Meters_RegisterNames(self):
        """ Método que retorna uma lista com o nome de todos os registros do medidor ativo"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(1), variant_pointer)
        return variant_pointer.contents.value

    def get_Meters_RegisterValues(self):
        """ Método que retorna uma lista com os valores do registros do medidor ativo"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(2), variant_pointer)
        return variant_pointer.contents.value

# Loads Interface

    # Loads int (int)
    def set_Loads_First(self):
        """ Ativa o primeiro elemento Load"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_Loads_Next(self):
        """ Ativa o próximo elemento Load"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_Load_Count(self):
        """ Método que retorna a quantidade de Loads"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(4), ctypes.c_int32(0))
        return resposta

    # Loads F (Float)
    def get_Loads_kV(self):
        """ Método que retorna o kV da carga ativada"""
        resposta = float(self.dssObj.DSSLoadsF(ctypes.c_int32(2), ctypes.c_double(0)))
        return resposta

    def set_Load_kW(self, kW):
        """ Método que seta o valor de kW da carga ativa"""
        resposta = float(self.dssObj.DSSLoadsF(ctypes.c_int32(1), ctypes.c_double(kW)))
        return resposta

    # Loads S (String)
    def get_Loads_Name(self):
        """ Método que retorna o nome da carga ativada"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def set_Loads_ActiveName(self, nome):
        """ Método que torna uma carga ativa pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def set_Loads_daily(self, nome):
        """ Método que seta o valor do parâmetro daily da carga ativa"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(5), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    # Loads V (Variant)
    def get_Loads_AllNames(self):
        """ Método que retorna uma lista com os nomes de todas as cargas do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.DSSLoadsV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

# Text Interface
    def text(self, comando):
        """ Método que envia um comando pela OpenDSSDirect.dll """
        resposta = ctypes.c_char_p(self.dssObj.DSSPut_Command(comando.encode('ascii')))
        return resposta.value

# Lines Interface

    # LinesI (int)
    def set_Lines_First(self):
        """ Ativa o primeiro elemento Line """
        resposta = self.dssObj.LinesI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_Lines_Next(self):
        """ Ativa o próximo elemento Line """
        resposta = self.dssObj.LinesI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_Lines_Count(self):
        """ Método que retorna a quantidade de Line """
        resposta = self.dssObj.LinesI(ctypes.c_int32(6), ctypes.c_int32(0))
        return resposta

    # LinesF (Float)
    # LinesS (String)
    def get_Lines_Name(self):
        """ Método que retorna o nome do line ativado"""
        resposta = ctypes.c_char_p(self.dssObj.LinesS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_Lines_Bus1(self):
        """ Método que retorna o nome da barra 1 do line ativado"""
        resposta = ctypes.c_char_p(self.dssObj.LinesS(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_Lines_Bus2(self):
        """ Método que retorna o nome da barra 1 do line ativado"""
        resposta = ctypes.c_char_p(self.dssObj.LinesS(ctypes.c_int32(4), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # LinesV (Variant)

# Capacitors Interface

    # CapacitorsI (int)
    # CapacitorsF (Float)
    # CapacitorsS (String)
    # CapacitorsV (Variant)

# Transformers Interface

    # TransformersI (int)
    def get_Transformers_NumWindings(self):
        """ Método que retorna a quantidade de enrolamentos do transformador ativado"""
        resposta = self.dssObj.TransformersI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_Transformers_First(self):
        """ Ativa o primeiro elemento Transformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(8), ctypes.c_int32(0))
        return resposta

    def set_Transformers_Next(self):
        """ Ativa o próximo elementoTransformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(9), ctypes.c_int32(0))
        return resposta

    def get_Transformers_Count(self):
        """ Método que retorna a quantidade de Transformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(10), ctypes.c_int32(0))
        return resposta

    # TransformersF (Float)
    def set_Transformers_kVA(self, kva):
        """ Método que seta a potência aparente do transformador ativo"""
        resposta = float(self.dssObj.TransformersF(ctypes.c_int32(11), ctypes.c_double(kva)))
        return resposta

    def get_Transformers_kVA(self):
        """ Método que retorna a potência aparente do transformador ativo"""
        resposta = float(self.dssObj.TransformersF(ctypes.c_int32(10), ctypes.c_double(0)))
        return resposta

    # TransformersS (String)
    def set_Transfomers_ActiveName(self, nome):
        """ Método que seta um transformador ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.TransformersS(ctypes.c_int32(3), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def get_Transformers_Name(self):
        """ Método que retorna o nome do transformer ativado"""
        resposta = ctypes.c_char_p(self.dssObj.TransformersS(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # TransformersV (Variant)
    def get_Transformers_AllNames(self):
        """ Método que retorna uma lista com os nomes de todos os Transformadores"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.TransformersV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

# CapControls Interface

    # CapControlsI (int)
    def set_CapControls_First(self):
        """ Ativa o primeiro elemento CapControl"""
        resposta = self.dssObj.CapControlsI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_CapControls_Next(self):
        """ Ativa o próximo elemento CapControl"""
        resposta = self.dssObj.CapControlsI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_CapControls_Count(self):
        """ Método que retorna a quantidade de CapControls"""
        resposta = self.dssObj.CapControlsI(ctypes.c_int32(8), ctypes.c_int32(0))
        return resposta
    # CapControlsF (Float)
    # CapControlsS (String)
    def get_CapControls_Capacitor(self):
        """ Método que retorna o nome do tcapacitor controlado pelo CapControl ativado"""
        resposta = ctypes.c_char_p(self.dssObj.CapControlsS(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')
    # CapControlsV (Variant)

# RegControls Interface

    # RegControlsI (int)
    def set_RegControls_First(self):
        """ Ativa o primeiro elemento RegControl"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_RegControls_Next(self):
        """ Ativa o próximo elemento RegControl"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_RegControls_Count(self):
        """ Método que retorna a quantidade de RegControls"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(12), ctypes.c_int32(0))
        return resposta

    def get_RegControls_TapWinding(self):
        """ Método que retorna o terminal do tap"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(2), ctypes.c_int32(0))
        return resposta

    # RegControlsF (Float)
    # RegControlsS (String)
    def get_RegControls_Transformer(self):
        """ Método que retorna o nome do transformador controlado pelo RegControl ativado"""
        resposta = ctypes.c_char_p(self.dssObj.RegControlsS(ctypes.c_int32(4), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_RegControls_Name(self):
        """ Método que retorna o nome do RegControl ativado"""
        resposta = ctypes.c_char_p(self.dssObj.RegControlsS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # RegControlsV (Variant)

# VSources Interface

    # VSourcesI (int)
    def set_Vsources_First(self):
        """ Ativa o primeiro elemento Vsource"""
        resposta = self.dssObj.VsourcesI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    # VSourcesF (Float)
    def get_Vsources_BasekV(self):
        """ Método que retorna a tensão de base do elemento circuit"""
        resposta = float(self.dssObj.VsourcesF(ctypes.c_int32(0), ctypes.c_double(0)))
        return resposta

    # VSourcesS (String)
    # VSourcesV (Variant)

# Topology Interface

    # TopologyI (int)
    def topology_BackwardBranch(self):
        """ Método que torna o ramo à montante ativo. Retorna o index do elemento se encontrou. Caso não haja mais, retorna 0"""
        resposta = int(self.dssObj.TopologyI(ctypes.c_int32(7), ctypes.c_int32(0)))
        return resposta

    # TopologyF (Float)
    # TopologyS (String)
    def set_Topology_ActiveBranchName(self, nome):
        """ Método que seta um ramo ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.TopologyS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def get_Topology_ActiveBranchName(self):
        """ Método que retorna o nome do ramo ativo"""
        resposta = ctypes.c_char_p(self.dssObj.TopologyS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    # TopologyV (Variant)
