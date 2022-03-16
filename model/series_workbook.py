import math
import os.path

import numpy as np
from openpyxl import load_workbook
from scipy import interpolate


class SeriesWorkbook:
    MAX_SERIES = 5

    def __init__(self):
        self.workbook_path = None
        self.workbook = None

    def open_workbook(self, workbook_path):
        self.workbook_path = workbook_path
        self.workbook = load_workbook(workbook_path)

        sheet_names = self.workbook.sheetnames[1:]
        print('Sheets: %s' % ' '.join(sheet_names))

    def process_workbook(self):
        for sheet in self.workbook.worksheets[1:]:
            print("Processing sheet: %s" % sheet.title)
            self.process_sheet(sheet)

        basename, extension = os.path.splitext(self.workbook_path)
        self.workbook.save(basename + '_avg' + extension)

    def process_sheet(self, sheet):
        series = []

        for n in range(0, self.MAX_SERIES):
            print("Processing series: %d" % n)
            series.append(self.create_series(sheet, n))

        average = self.create_average(series)
        self.save_series(sheet, average)

    def create_series(self, sheet, series_n):
        base_column = self.get_series_base_column(series_n)
        base_row = self.get_series_base_row()

        series = Series()
        series.t = self.load_single_series(sheet, base_column, base_row, sheet.max_row)
        series.x = self.load_single_series(sheet, base_column + 1, base_row, sheet.max_row)
        series.y = self.load_single_series(sheet, base_column + 2, base_row, sheet.max_row)

        series.xt = interpolate.interp1d(series.t, series.x, bounds_error=False)
        series.yt = interpolate.interp1d(series.t, series.y, bounds_error=False)

        return series

    def create_average(self, series):
        average = Series()
        t_start = series[0].t[0]
        t_end = series[0].t[-1]
        samples = 0

        for s in series:
            t_start = min(t_start, s.t[0])
            t_end = max(t_end, s.t[-1])
            samples = max(samples, len(s.t))

        for t in np.linspace(t_start, t_end, num=samples):
            x_values = [s.xt(t) for s in series if not math.isnan(s.xt(t))]
            y_values = [s.yt(t) for s in series if not math.isnan(s.yt(t))]

            x_avg = sum(x_values) / len(x_values)
            y_avg = sum(y_values) / len(y_values)

            average.t.append(t)
            average.x.append(x_avg)
            average.y.append(y_avg)

        return average

    def save_series(self, sheet, series, series_n=5):
        base_column = self.get_series_base_column(series_n)
        base_row = self.get_series_base_row()

        for n in range(len(series.t)):
            sheet.cell(row=base_row + n, column=base_column).value = series.t[n]
            sheet.cell(row=base_row + n, column=base_column + 1).value = series.x[n]
            sheet.cell(row=base_row + n, column=base_column + 2).value = series.y[n]

    def load_single_series(self, sheet, column, row_start, row_end):
        values = []

        for row in range(row_start, row_end):
            value = sheet.cell(column=column, row=row).value
            values.append(value)

        values = [v for v in values if v is not None]
        return np.array(values)

    def get_series_base_column(self, series_n):
        return 2 + series_n * 3

    def get_series_base_row(self):
        return 5


class Series:
    def __init__(self):
        self.t = []
        self.x = []
        self.y = []
        self.xt = None
        self.yt = None
