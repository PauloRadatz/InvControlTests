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




if __name__ == '__main__':
    sns.set(style="ticks")
    sns.set_context("notebook")
    sns.set_style("whitegrid")


    feeder = "123Bus"
    feeder = "creelman"

    if feeder == "123Bus":
        output1 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus"
        output2 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\123Bus_modified4"
    if feeder == "creelman":
        output1 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman"
        output2 = r"G:\Drives compartilhados\Celso-Paulo\EPRI\2019\AgnosticInvControlModel\Task1\Tests\PVSystem\ConvergenceTests\Output\Creelman_modified4"

    case_dic = {"c1": r"\1-50_10_Q",        "c2": r"\2-50_10_P",        "c3": r"\3-50_10_All",
                "c4": r"\4-50_50_Q",        "c5": r"\5-50_50_P",        "c6": r"\6-50_50_All",
                "c7": r"\7-50_100_Q",       "c8": r"\8-50_100_P",       "c9": r"\9-50_100_All",
                "c10": r"\10-100_10_Q",     "c11": r"\11-100_10_P",     "c12": r"\12-100_10_All",
                "c13": r"\13-100_50_Q",     "c14": r"\14-100_50_P",     "c15": r"\15-100_50_All",
                "c16": r"\16-100_100_Q",    "c17": r"\17-100_100_P",    "c18": r"\18-100_100_All",
                "c19": r"\19-150_10_Q",     "c20": r"\20-150_10_P",     "c21": r"\21-150_10_All",
                "c22": r"\22-150_50_Q",     "c23": r"\23-150_50_P",     "c24": r"\24-150_50_All",
                "c25": r"\25-150_100_Q",    "c26": r"\26-150_100_P",    "c27": r"\27-150_100_All"}

    #case_dic = ["c25": r"\25-150_100_Q"]

    a = Main(case_dic, output1, output2)
    a.plot_individual()
    a.plot_diff()

