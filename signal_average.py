import sys


def main():
    if len(sys.argv) < 1:
        print('Signal Average: no workbook provided')

    workbook_path = sys.argv[1]

    print('Signal Average')
    print('Processing workbook: %s' % workbook_path)


if __name__ == '__main__':
    main()
