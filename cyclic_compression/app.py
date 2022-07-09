import matplotlib.pyplot as plt
from scipy.signal import find_peaks

from common.workbook_wrapper import open_workbook, save_with_suffix, Cell
from data.config import config


def main():
    print('Cyclic compression calculations')

    for workbook_name, sheets in config.items():
        workbook = open_workbook('data/' + workbook_name)
        print('Processing workbook: ' + workbook.workbook_path)

        for sheet_name, sheet in sheets.items():
            print("Processing sheet: " + sheet_name)
            workbook.set_sheet('3-1-3')

            for dataset_n, dataset in sheet.items():
                print('Processing dataset #{}: {}'.format(dataset_n, str(list(dataset.values()))))
                t_series = workbook.load_single_series(Cell('B5'))
                e_series = workbook.load_single_series(Cell('C5'))
                minima, _ = find_peaks(-e_series)
                plt.plot(t_series, e_series)
                plt.plot(t_series[minima], e_series[minima], '.')
                plt.show()

        save_with_suffix(workbook, "new")


if __name__ == "__main__":
    main()
