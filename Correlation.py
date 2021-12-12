import pandas as pd
import math
import os
import ctypes

if os.path.exists('Data.xlsx'):
    pass
else:
    ctypes.windll.user32.MessageBoxW(0, 'Please add Data.xlsx to the current directory',
                                     'Data File not found', 1)


class correl:
    """Create a set of ordered data in Data.xlsx in the same directory,
    this will then iterate through the columns and calculate the correlation between them"""

    def __init__(self):
        self.location = 'Data.xlsx'
        self.analysis = dict()
        self.len_check()

    def len_check(self):
        self.check = 0
        self.df = pd.read_excel(self.location)
        if self.df.isnull().values.any():
            self.check = 1
        else:
            pass
        if self.check == 0:
            self.excel_import()
        else:
            ctypes.windll.user32.MessageBoxW(0, 'Please ensure all columns are the same length',
                                             'Data length not consistent', 1)

    def excel_import(self):
        self.col = list(self.df.columns)
        for name in self.col:
            self.colnames = set([name]).symmetric_difference(set(self.col))
            for compare in self.colnames:
                pearson = self.NXY(list(self.df[name]), list(self.df[compare]))
                self.analysis[name + ' & ' + compare] = pearson

        self.output = pd.DataFrame(self.analysis.items(), columns=['Variables', 'Correlation'])
        self.output.to_excel('Output.xlsx', index=False)

    def NXY(self, x, y):
        self.x = x
        self.y = y
        self.nx = len(self.x)
        self.ny = len(self.y)
        self.t1 = int()
        for a, b in zip(self.x, self.y):
            self.t1 += (a * b)
        self.t1 = self.t1 * self.nx
        self.t2 = sum(self.x) * sum(self.y)
        self.top = self.t1 - self.t2

        self.b1 = float()
        self.b3 = float()
        for a, b in zip(self.x, self.y):
            self.b1 += a ** 2
            self.b3 += b ** 2
        self.b1 = self.nx * self.b1
        self.b3 = self.ny * self.b3

        self.b2 = sum(self.x) ** 2
        self.b4 = sum(self.y) ** 2
        self.bottom = math.sqrt((self.b1 - self.b2) * (self.b3 - self.b4))
        self.result = self.top / self.bottom

        return self.result


correl()
