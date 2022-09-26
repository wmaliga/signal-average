import matplotlib.pyplot as plt
import numpy
from scipy.signal import find_peaks

from common.signal_utils import signal_average
from common.workbook_wrapper import open_workbook, save_with_suffix, Cell
from data.config import config


def main():
    print('Cyclic compression calculations')
    save_steps = False

    for workbook_name, sheets in config.items():
        workbook = open_workbook(f'data/{workbook_name}')
        print(f'Processing workbook: {workbook.workbook_path}')

        for sheet_name, sheet in sheets.items():
            print(f'Processing sheet: {sheet_name}')
            workbook.set_sheet(sheet_name)

            for dataset_n, dataset in sheet.items():
                print(f'Processing dataset #{dataset_n}: {str(list(dataset.values()))}')
                t_series = workbook.load_single_series(Cell(dataset['t']))
                e_series = workbook.load_single_series(Cell(dataset['E']))
                l_series = workbook.load_single_series(Cell(dataset['L']))
                minima, _ = find_peaks(-e_series)
                minima = minima[4::5]
                if save_steps:
                    plt.plot(t_series, e_series)
                    plt.plot(t_series[minima], e_series[minima], '.')
                    plt.savefig('out/step1.png')
                    plt.show()

                t_series_list = numpy.split(t_series, minima)
                e_series_list = numpy.split(e_series, minima)
                l_series_list = numpy.split(l_series, minima)

                cycles = len(t_series_list)
                print(f'Split into {cycles} cycles')

                t_series_list = t_series_list[1:]
                e_series_list = e_series_list[1:]
                l_series_list = l_series_list[1:]

                t_series_list = [t_series - t_series[0] for t_series in t_series_list]

                if save_steps:
                    for i in range(len(t_series_list)):
                        plt.plot(t_series_list[i], e_series_list[i])
                    plt.savefig('out/step2e.png')
                    plt.show()

                    for i in range(len(t_series_list)):
                        plt.plot(t_series_list[i], l_series_list[i])
                    plt.savefig('out/step2l.png')
                    plt.show()

                t_avg_values, e_avg_values, e_dev_values = signal_average(t_series_list, e_series_list)
                t_avg_values, l_avg_values, l_dev_values = signal_average(t_series_list, l_series_list)

                if save_steps:
                    plt.plot(t_avg_values, e_avg_values)
                    plt.savefig('out/step3e.png')
                    plt.show()

                    plt.plot(t_avg_values, l_avg_values)
                    plt.savefig('out/step3l.png')
                    plt.show()

                sheet_name_avg = f'{sheet_name}_avg'
                workbook.create_sheet(sheet_name_avg)
                workbook.write_single_series(Cell(dataset['t']), t_avg_values, override_sheet=sheet_name_avg)
                workbook.write_single_series(Cell(dataset['E']), e_avg_values, override_sheet=sheet_name_avg)
                workbook.write_single_series(Cell(dataset['L']), l_avg_values, override_sheet=sheet_name_avg)

                save_steps = False

        save_with_suffix(workbook, "new")


if __name__ == "__main__":
    main()
