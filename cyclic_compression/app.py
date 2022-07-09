import matplotlib.pyplot as plt
import numpy
from scipy.signal import find_peaks

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

                #for i in range(1, cycles):
                #    plt.plot(t_series_list[i] - t_series_list[i][0], e_series_list[i])

                for i in range(1, cycles):
                    plt.plot(t_series_list[i] - t_series_list[i][0], l_series_list[i])

                plt.show()

        save_with_suffix(workbook, "new")


if __name__ == "__main__":
    main()
