import sys

from model.series_workbook import SeriesWorkbook


def main():
    if len(sys.argv) < 1:
        print('Signal Average: no workbook provided')

    workbook_path = sys.argv[1]

    print('Signal Average')
    print('Processing workbook: %s' % workbook_path)

    signals = SeriesWorkbook()
    signals.open_workbook(workbook_path)
    signals.process_workbook()


if __name__ == '__main__':
    main()
