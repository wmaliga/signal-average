import sys

from data.config import data
from model.series_workbook import SeriesWorkbook


def main():
    print('Signal Average')
    workbook = SeriesWorkbook()

    for workbook_basename, sheets in data.items():
        workbook_path = 'data/{}.xlsx'.format(workbook_basename)
        print('Workbook: ' + workbook_path)
        workbook.open_workbook(workbook_path)

        for sheet_name, sheet in sheets.items():
            print('Sheet: ' + sheet_name)
            workbook.set_sheet(sheet_name)

            for series_n, series in sheet.items():
                print('Series: ' + str(series_n))

                for series_letter in series:
                    if len(series_letter) == 1:
                        print('{}: {} -> {}'.format(series_letter, ','.join(series[series_letter]), series[series_letter + '_avg']))

                workbook.process_series(series)


if __name__ == '__main__':
    main()
