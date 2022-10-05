import matplotlib.pyplot as plt
import numpy
from scipy.signal import find_peaks

from common.signal_utils import chunks, signal_average
from common.workbook_wrapper import open_workbook, save_with_suffix, Cell
from data.config import config

colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k']


def main():
    print('Cyclic compression calculations')

    for workbook_name, sheets in config.items():
        workbook = open_workbook(f'data/{workbook_name}')
        print(f'Processing workbook: {workbook.workbook_path}')

        for sheet_name, sheet in sheets.items():
            print(f'Processing sheet: {sheet_name}')
            workbook.set_sheet(sheet_name)

            for dataset_n, dataset in sheet.items():
                print(f'Processing dataset #{dataset_n}: {str(list(dataset.values()))}')
                plot_steps = dataset.get('plot', False)

                t_series = workbook.load_single_series(Cell(dataset['t']))
                e_series = workbook.load_single_series(Cell(dataset['E']))
                l_series = workbook.load_single_series(Cell(dataset['L']))

                minima, _ = find_peaks(-e_series, distance=50, width=5)

                if plot_steps:
                    plt.plot(t_series, e_series)
                    plt.plot(t_series[minima], e_series[minima], '.')
                    plt.show()

                t_series_list = numpy.split(t_series, minima)
                e_series_list = numpy.split(e_series, minima)
                l_series_list = numpy.split(l_series, minima)

                cycles = len(t_series_list)
                print(f'Split into {cycles} cycles')

                t_series_groups = list(chunks(t_series_list, 5))
                e_series_groups = list(chunks(e_series_list, 5))
                l_series_groups = list(chunks(l_series_list, 5))

                if plot_steps:
                    for i in range(len(t_series_groups)):
                        color = colors[i % len(colors)]
                        for j in range(len(t_series_groups[i])):
                            plt.plot(t_series_groups[i][j], e_series_groups[i][j], color)
                    plt.show()

                    for i in range(len(t_series_groups)):
                        color = colors[i % len(colors)]
                        for j in range(len(t_series_groups[i])):
                            plt.plot(t_series_groups[i][j], l_series_groups[i][j], color)
                    plt.show()

                t_series_groups = t_series_groups[1:]
                e_series_groups = e_series_groups[1:]
                l_series_groups = l_series_groups[1:]

                for i in range(len(t_series_groups)):
                    group_start_t = t_series_groups[i][0][0]
                    t_series_groups[i] = [t_series - t_series[0] + group_start_t for t_series in t_series_groups[i]]

                if plot_steps:
                    for i in range(len(t_series_groups)):
                        for j in range(len(t_series_groups[i])):
                            plt.plot(t_series_groups[i][j], e_series_groups[i][j])
                    plt.show()

                    for i in range(len(t_series_groups)):
                        for j in range(len(t_series_groups[i])):
                            plt.plot(t_series_groups[i][j], l_series_groups[i][j])
                    plt.show()

                t_avg_groups = []
                e_avg_groups = []
                e_dev_groups = []
                l_avg_groups = []
                l_dev_groups = []

                for i in range(len(t_series_groups)):
                    t_avg, e_avg, e_dev = signal_average(t_series_groups[i], e_series_groups[i])
                    t_avg_groups.append(t_avg)
                    e_avg_groups.append(e_avg)
                    e_dev_groups.append(e_dev)
                    t_avg, l_avg, l_dev = signal_average(t_series_groups[i], l_series_groups[i])
                    l_avg_groups.append(l_avg)
                    l_dev_groups.append(l_dev)

                prev_t_end = 0

                for i in range(len(t_avg_groups)):
                    t_avg_groups[i] = [t - t_avg_groups[i][0] + prev_t_end for t in t_avg_groups[i]]
                    prev_t_end = t_avg_groups[i][-1]

                if plot_steps:
                    for i in range(len(t_avg_groups)):
                        plt.plot(t_avg_groups[i], e_avg_groups[i])
                    plt.show()

                    for i in range(len(t_avg_groups)):
                        plt.plot(t_avg_groups[i], l_avg_groups[i])
                    plt.show()

                t_avg_values = []
                e_avg_values = []
                e_dev_values = []
                l_avg_values = []
                l_dev_values = []

                for i in range(len(t_avg_groups)):
                    t_avg_values.extend(t_avg_groups[i])
                    e_avg_values.extend(e_avg_groups[i])
                    e_dev_values.extend(e_dev_groups[i])
                    l_avg_values.extend(l_avg_groups[i])
                    l_dev_values.extend(l_dev_groups[i])

                if plot_steps:
                    plt.plot(t_avg_values, e_avg_values)
                    plt.show()

                    plt.plot(t_avg_values, l_avg_values)
                    plt.show()

                sheet_name_avg = f'{sheet_name}_avg'
                workbook.create_sheet(sheet_name_avg)
                workbook.write_single_series(Cell(dataset['t']), t_avg_values, override_sheet=sheet_name_avg)
                workbook.write_single_series(Cell(dataset['E']), e_avg_values, override_sheet=sheet_name_avg)
                workbook.write_single_series(Cell(dataset['L']), l_avg_values, override_sheet=sheet_name_avg)

        save_with_suffix(workbook, "new")


if __name__ == "__main__":
    main()
