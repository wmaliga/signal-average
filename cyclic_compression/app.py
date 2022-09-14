import matplotlib.pyplot as plt
import numpy
from scipy.signal import find_peaks

from common.signal_utils import signal_average
from common.workbook_wrapper import open_workbook, save_with_suffix, Cell
from data.config import config


def main():
    print('Cyclic compression calculations')

    for workbook_name, sheets in config.items():
        workbook = open_workbook(f'data/{workbook_name}')
        print(f'Processing workbook: {workbook.workbook_path}')

        for sheet_name, sheet in sheets.items():
            print(f'Processing sheet: {sheet_name}')
            workbook.set_sheet('3-1-3')

            for dataset_n, dataset in sheet.items():
                print(f'Processing dataset #{dataset_n}: {str(list(dataset.values()))}')
                t_series = workbook.load_single_series(Cell(dataset['t']))
                e_series = workbook.load_single_series(Cell(dataset['E']))
                l_series = workbook.load_single_series(Cell(dataset['L']))
                minima, _ = find_peaks(-e_series)
                # plt.plot(t_series, e_series)
                # plt.plot(t_series[minima], e_series[minima], '.')
                # plt.show()
                t_series_list = numpy.split(t_series, minima)
                e_series_list = numpy.split(e_series, minima)
                l_series_list = numpy.split(l_series, minima)

                cycles = len(t_series_list)
                print(f'Split into {cycles} cycles')

                t_series_list = t_series_list[1:]
                e_series_list = e_series_list[1:]
                l_series_list = l_series_list[1:]

                t_series_list = [t_series - t_series[0] for t_series in t_series_list]

                # for i in range(len(t_series_list)):
                #     plt.plot(t_series_list[i], e_series_list[i])

                # for i in range(len(t_series_list)):
                #     plt.plot(t_series_list[i], l_series_list[i])

                # plt.show()

                t_avg_values, e_avg_values, e_dev_values = signal_average(t_series_list, e_series_list)
                t_avg_values, l_avg_values, l_dev_values = signal_average(t_series_list, l_series_list)

                # plt.plot(t_avg_values, e_avg_values)
                # plt.plot(t_avg_values, l_avg_values)
                # plt.show()

                workbook.create_sheet('3-1-3_avg')
                workbook.write_single_series(Cell('A1'), t_avg_values, override_sheet='3-1-3_avg')
                workbook.write_single_series(Cell('B1'), e_avg_values, override_sheet='3-1-3_avg')
                workbook.write_single_series(Cell('C1'), l_avg_values, override_sheet='3-1-3_avg')

        save_with_suffix(workbook, "new")


if __name__ == "__main__":
    main()
