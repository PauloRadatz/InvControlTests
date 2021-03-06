###  # -*- coding: iso-8859-15 -*-
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.lines as mlines
from matplotlib.font_manager import FontProperties
import os
import sqlite3

fontPx = FontProperties()
fontP = FontProperties()
fontPm = FontProperties()
fontPx.set_size('x-small')
fontP.set_size('small')
fontPm.set_size('medium')


class Main(object):

    def __init__(self, case_dic, output1, output2):

        self.case_dic = case_dic
        self.output1 = output1
        self.output2 = output2

        self.output_list = [self.output1, self.output2]

        self.df_dic = {}

        #self.df_dic_diff = {}

        self.df_diff_list = []

        self.flag_dif = False

    def create_df_dic(self, output, columns=None):

        if columns:
            for case in self.case_dic:
                penetration = self.case_dic[case].split("_")[-3].split("-")[-1]
                buses = self.case_dic[case].split("_")[-2]
                type = self.case_dic[case].split("_")[-1]
                self.df_dic[case] = pd.read_csv(output + self.case_dic[case] + "\conditions_results.csv")[columns]
                self.df_dic[case]["penetration"] = penetration
                self.df_dic[case]["buses"] = buses
                self.df_dic[case]["type"] = type
        else:
            for case in self.case_dic:
                self.df_dic[case] = pd.read_csv(output + self.case_dic[case] + "\conditions_results.csv")

    def rename_columns_df(self, columns):

        for df in self.df_dic:
            self.df_dic[df].rename(columns, axis=1, inplace=True, copy=True)

    def create_total_df(self):

        self.total_df = pd.concat(self.df_dic)

    def plot_factorplot(self, data, x, y, hue, col, title, name, output):
        # https://www.programcreek.com/python/example/96202/seaborn.factorplot
        g = sns.catplot(x=x, y=y, hue=hue, data=data, kind="bar", col=col, palette="muted", legend=True,
                           ci=None, legend_out=True, sharex=True, sharey=False, hue_order=["10", "50", "100"], col_order=["50", "100", "150"], row_order=None)
        g.set_axis_labels(x, y)
        if title:
            #g.set_titles(title)
            g.fig.suptitle(title)
            g.fig.subplots_adjust(top=0.9)

        if not self.flag_dif:
            g.set(ylim=(0, 100))
            g.set(yticks=[0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100])
        else:
            g.set(ylim=(-100, 100))
            g.set(yticks=[-100, -90, -80, -70, -60, -50, -40, -30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100])

        if feeder == "123Bus":
            file_folder = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus\Figure" + "\\" + name + ".png"

        if feeder == "creelman":
            file_folder = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman_modified4\Figure" + "\\" + name + ".png"
            #file_folder = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman\Figure" + "\\" + name + ".png"

        plt.savefig(
            output + "\\" + "Figure" + "\\" + name + ".png",
            dpi=None, facecolor='w', edgecolor='w',
            orientation='portrait', papertype=None, format=None,
            transparent=False, bbox_inches="tight", pad_inches=0.1,
            frameon=None)

        #plt.show()

    def plot_factorplot_All(self, data, x, y, hue, col, row, title, name, output):

        g = sns.catplot(data=data, x=x, y=y, row=row, col=col, hue=hue, palette="muted", kind="bar", legend=True,
                           ci=None, legend_out=True, aspect=1, hue_order=["10", "50", "100"], row_order=["50", "100", "150"], sharey=False)
        g.fig.suptitle(title)
        g.fig.subplots_adjust(top=.9)
        #g.fig.set_size_inches(6, 5, forward=True)
        #g.set_ylabels("Scenarios")

        if not self.flag_dif:
            g.set(ylim=(0, 100))
            g.set(yticks=[0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100])
        else:
            g.set(ylim=(-100, 100))
            g.set(yticks=[-100, -90, -80, -70, -60, -50, -40, -30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100])

        if feeder == "123Bus":
            file_folder = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus\Figure" + "\\" + name + ".png"

        if feeder == "creelman":
            file_folder = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman_modified4\Figure" + "\\" + name + ".png"
            #file_folder = r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman\Figure" + "\\" + name + ".png"

        # r"G:\Drives de equipe\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus\Figure"
        plt.savefig(output + "\\" + "Figure" + "\\" + name + ".png", dpi=None, facecolor='w', edgecolor='w',
                    orientation='portrait', papertype=None, format=None,
                    transparent=False, bbox_inches="tight", pad_inches=0.1,
                    frameon=None)


        #plt.show()

    def plot_All(self, output):

        x = "DeltaQ_factor"
        y = "Scenarios without Issue (%)"
        col = "DeltaP_factor"
        hue = "Three-phase buses (%)"
        row = "Penetration (%)"
        title = self.feeder + " with All functions"
        data = self.total_df[self.total_df["type"] == "All"]

        self.plot_factorplot_All(data, x, y, hue, col, row, title, "S-All", output)

        y = "mean - Num. Control Iteration"
        self.plot_factorplot_All(data, x, y, hue, col, row, title, "C-All", output)

    def plot_Q(self, output):
        hue_list = ["Three-phase buses (%)", "Penetration (%)"]

        # for hue in hue_list:
        #     self.plot_factorplot(x, y, hue)

        x = "DeltaQ_factor"
        y = "Scenarios without Issue (%)"
        hue = "Three-phase buses (%)"
        col = "Penetration (%)"
        title = self.feeder + " with Q functions"
        data = self.total_df[self.total_df["type"] == "Q"]
        self.plot_factorplot(data, x, y, hue, col, title, "S-Q", output)

        y = "mean - Num. Control Iteration"
        self.plot_factorplot(data, x, y, hue, col, title, "C-Q", output)

    def plot_P(self, output):
        x = "DeltaP_factor"
        y = "Scenarios without Issue (%)"
        # hue_list = ["Three-phase buses (%)", "Penetration (%)"]
        #
        # for hue in hue_list:
        #     self.plot_factorplot(x, y, hue)

        hue = "Three-phase buses (%)"
        col = "Penetration (%)"
        title = self.feeder + " with P functions"
        data = self.total_df[self.total_df["type"] == "P"]
        self.plot_factorplot(data, x, y, hue, col, title, "S-P", output)

        y = "mean - Num. Control Iteration"
        self.plot_factorplot(data, x, y, hue, col, title, "C-P", output)

    def plot_All_diff(self, output):

        x = "DeltaQ_factor"
        col = "DeltaP_factor"
        hue = "Three-phase buses (%)"
        row = "Penetration (%)"
        title = self.feeder + " with All functions"
        data = self.df_diff_total[self.df_diff_total["type"] == "All"]

        y = "Error of mean - Num. Control Iteration"
        self.plot_factorplot_All(data, x, y, hue, col, row, title, "C-All-Error", output)

    def plot_Q_diff(self, output):

        x = "DeltaQ_factor"
        hue = "Three-phase buses (%)"
        col = "Penetration (%)"
        title = self.feeder + " with Q functions"
        data = self.df_diff_total[self.df_diff_total["type"] == "Q"]

        y = "Error of mean - Num. Control Iteration"
        self.plot_factorplot(data, x, y, hue, col, title, "C-Q-Error", output)

    def plot_P_diff(self, output):

        x = "DeltaP_factor"
        hue = "Three-phase buses (%)"
        col = "Penetration (%)"
        title = self.feeder + " with P functions"
        data = self.df_diff_total[self.df_diff_total["type"] == "P"]

        y = "Error of mean - Num. Control Iteration"
        self.plot_factorplot(data, x, y, hue, col, title, "C-P-Error", output)

    def plot_individual(self):

        if feeder == "123Bus":
            self.feeder = "123Bus"
        if feeder == "creelman":
            self.feeder = "683"

        columns = ["Scenarios without Issue (%)", "mean - # Control Iteration", "DeltaP_factor", "DeltaQ_factor"]
        columns_renamed = {"Scenarios without Issue (%)": "Scenarios without Issue", "mean - # Control Iteration": "mean Control Iteration"}

        for output in self.output_list:
            self.create_df_dic(output, columns)
            self.rename_columns_df(columns_renamed)

            self.create_total_df()

            #self.df_dic_diff[output.split("\\")[-1]] = self.total_df
            self.df_diff_list.append(self.total_df)

            columns_renamed_total = {"Scenarios without Issue": "Scenarios without Issue (%)", "mean Control Iteration": "mean - Num. Control Iteration",
                           "penetration": "Penetration (%)", "buses": "Three-phase buses (%)"}

            self.total_df.rename(columns_renamed_total, axis=1, inplace=True, copy=True)

            self.plot_All(output)

            self.plot_Q(output)
            self.plot_P(output)

    def plot_diff(self):

        self.flag_dif = True
        df = self.df_diff_list[0]

        sr_base = self.df_diff_list[0]["mean - Num. Control Iteration"]
        sr = self.df_diff_list[1]["mean - Num. Control Iteration"]

        df["mean - Num. Control Iteration New"] = sr
        #df["Error of mean - Num. Control Iteration"] = 100.0 * np.abs(sr - sr_base) / sr

        df["Error of mean - Num. Control Iteration"] = 100.0 * (sr_base - sr) / sr

        self.df_diff_total = df

        self.plot_All_diff(self.output_list[-1])
        self.plot_Q_diff(self.output_list[-1])
        self.plot_P_diff(self.output_list[-1])
        print "here"

class MainQSTS(object):

    def __init__(self, case_dic, output_dic, drc):

        self.drc = drc
        self.case_order = ["c" + str(i + 1) for i in range(27)]
        if self.drc:
            self.columns_list = ["Update", "UpdateP", "Not Update", "0.4", "0.2", "0.1", "0.05"]
        else:
            self.columns_list = ["Update", "UpdateP", "Not Update", "0.2", "0.1", "0.05"]

        self.case_dic = case_dic
        self.output_dic = output_dic

        self.create_case_order()
        self.create_dfs()



    def create_case_order(self):
        penetration = []
        buses = []
        type = []
        case_list = []

        for case in self.case_dic:
            case_list.append(case)
            penetration.append(self.case_dic[case].split("_")[-3].split("-")[-1])
            buses.append(self.case_dic[case].split("_")[-2])
            type.append(self.case_dic[case].split("_")[-1])

        dic = {"case": case_list, "penetration": penetration, "buses": buses, "type": type}

        self.df_cases = pd.DataFrame(dic).set_index("case").reindex(self.case_order)[["penetration", "buses", "type"]]


    def create_dfs(self):

        case_name = []
        sr_scenarios_dic = {}
        sr_iterations_dic = {}

        df_scenarios_dic = {}
        df_iterations_dic = {}

        for key, output in self.output_dic.items():
            for case in self.case_dic:
                sr_scenarios_dic[case] = pd.read_csv(output + self.case_dic[case] + "\conditions_results.csv")[
                    "Scenarios without Issue (%)"].rename("Scenarios without Issue")
                sr_iterations_dic[case] = pd.read_csv(output + self.case_dic[case] + "\conditions_results.csv")[
                    "mean - # Control Iteration"].rename("Control Iteration Mean")

            df_scenarios_dic[key] = pd.concat(sr_scenarios_dic, axis=0).reset_index()
            df_scenarios_dic[key].set_index("level_0", inplace=True)
            del df_scenarios_dic[key]['level_1']
            df_scenarios_dic[key] = df_scenarios_dic[key]["Scenarios without Issue"].reindex(index=self.case_order)

            df_iterations_dic[key] = pd.concat(sr_iterations_dic, axis=0).reset_index()
            df_iterations_dic[key].set_index("level_0", inplace=True)
            del df_iterations_dic[key]['level_1']
            df_iterations_dic[key] = df_iterations_dic[key]["Control Iteration Mean"].reindex(index=self.case_order)

        self.df_scenarios = pd.concat(df_scenarios_dic, axis=1)[self.columns_list]
        self.df_iterations = pd.concat(df_iterations_dic, axis=1)[self.columns_list]

        self.df_scenarios[["penetration", "buses", "type"]] = self.df_cases
        self.df_iterations[["penetration", "buses", "type"]] = self.df_cases
        #self.df_iterations["case"] = self.df_iterations.index

        if self.drc:
            self.df_scenarios[["penetration", "buses", "type"] + self.columns_list].to_csv(r'G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Figure\df_scenarios.csv')
        else:
            self.df_scenarios[["penetration", "buses", "type"] + self.columns_list].to_csv(
                r'G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Figure\df_scenarios.csv')

        print "here"

    def plots(self, output):

        for function in ["Q", "P", "All"]:
            data_iteration = self.df_iterations[self.df_iterations["type"] == function][self.columns_list]
            self.box_plot(data_iteration, output, 'C_Box_' + function, function + " functions", "mean - Num. Control Iteration")

            #data_scenario = self.df_scenarios[self.df_scenarios["type"] == function][self.columns_list]
            #self.box_plot(data_scenario, output, 'S_Box_' + function, function + " functions", "Scenarios without Issue")

        self.bar_plot(100 - self.df_scenarios[self.columns_list], output, 'S_Bar', None, "Scenarios WITH Issue (%)")
        self.bar_plot(self.df_iterations[self.columns_list], output, 'C_Bar', None, "mean - Num. Control Iteration")

        # for function in ["Q", "P", "All"]:
        #     data_iteration = self.df_iterations[self.df_iterations["type"] == function][self.columns_list]
        #     self.bar_plot(data_iteration, output, 'C_Bar_' + function, function + " functions", "Control Iteration")
        #
        #     data_scenario = self.df_scenarios[self.df_scenarios["type"] == function][self.columns_list]
        #     self.bar_plot(data_scenario, output, 'S_Bar_' + function, function + " functions", "Scenarios without Issue")

    def box_plot(self, data, output, name, title, ylabel):
        plt.clf()
        sns.boxplot(data=data)
        #sns.stripplot(data=data, jitter=True, color='k')

        if title:
            # g.set_titles(title)
            plt.title(title)

        plt.ylabel(ylabel)

        plt.savefig(
            output + "\\" + "Figure" + "\\" + name + ".png",
            dpi=None, facecolor='w', edgecolor='w',
            orientation='portrait', papertype=None, format=None,
            transparent=False, bbox_inches="tight", pad_inches=0.1,
            frameon=None)

        #plt.show()

    def bar_plot(self, data, output, name, title, ylabel):
        #sns.barplot(data=data)
        #sns.stripplot(data=data, jitter=True, color='k')
        plt.clf()
        data.plot(kind="bar")

        if title:
            # g.set_titles(title)
            plt.title(title)

        plt.ylabel(ylabel)
        plt.legend(bbox_to_anchor=(1.1, 1.05))

        plt.savefig(
            output + "\\" + "Figure" + "\\" + name + ".png",
            dpi=None, facecolor='w', edgecolor='w',
            orientation='portrait', papertype=None, format=None,
            transparent=False, bbox_inches="tight", pad_inches=0.1,
            frameon=None)

        #plt.show()

    def box_plot_scenario(self, data):
        sns.boxplot(data=data)

        plt.ylabel("Scenarios without Issue")

        plt.show()

if __name__ == '__main__':
    sns.set(style="ticks")
    sns.set_context("notebook")
    sns.set_style("whitegrid")

    mode = 'qsts'


    feeder = "123Bus"
    feeder = "creelman"

    case_dic = {"c1": r"\1-50_10_Q", "c2": r"\2-50_10_P", "c3": r"\3-50_10_All",
                "c4": r"\4-50_50_Q", "c5": r"\5-50_50_P", "c6": r"\6-50_50_All",
                "c7": r"\7-50_100_Q", "c8": r"\8-50_100_P", "c9": r"\9-50_100_All",
                "c10": r"\10-100_10_Q", "c11": r"\11-100_10_P", "c12": r"\12-100_10_All",
                "c13": r"\13-100_50_Q", "c14": r"\14-100_50_P", "c15": r"\15-100_50_All",
                "c16": r"\16-100_100_Q", "c17": r"\17-100_100_P", "c18": r"\18-100_100_All",
                "c19": r"\19-150_10_Q", "c20": r"\20-150_10_P", "c21": r"\21-150_10_All",
                "c22": r"\22-150_50_Q", "c23": r"\23-150_50_P", "c24": r"\24-150_50_All",
                "c25": r"\25-150_100_Q", "c26": r"\26-150_100_P", "c27": r"\27-150_100_All"}

    if mode == 'snap':

        if feeder == "123Bus":
            output1 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus"
            output2 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus_modified4"
        if feeder == "creelman":
            output1 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman"
            output2 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman_modified4"


        #case_dic = ["c25": r"\25-150_100_Q"]

        a = Main(case_dic, output1, output2)
        a.plot_individual()
        a.plot_diff()

    else:

        drc = True

        if drc:
            output_update = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_Update"
            output_updatep = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_UpdateP"
            output_notupdate = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_NotUpdate"
            output_005 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_005"
            output_01 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_01"
            output_02 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_02"
            output_04 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC\Creelman_04"
            output_dic = {'Update': output_update, 'UpdateP': output_updatep, 'Not Update': output_notupdate,
                          '0.4': output_04, '0.2': output_02, '0.1': output_01, '0.05': output_005}

        else:
            output_update = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_Update"
            output_updatep = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_UpdateP"
            output_notupdate = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_NotUpdate"
            output_005 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_005"
            output_01 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_01"
            output_02 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\Creelman_02"
            output_dic = {'Update': output_update, 'UpdateP': output_updatep, 'Not Update': output_notupdate,
                          '0.2': output_02, '0.1': output_01, '0.05': output_005}

        a = MainQSTS(case_dic, output_dic, drc)
        if drc:
            a.plots(r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS\No_DRC")
        else:
            a.plots(r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output_QSTS")

        print "here"