from data.config import data
from model.data_set import data_set_from_map, validate_data_set
from model.workbook_wrapper import WorkbookWrapper


def main():
    print('Signal Average')
    wrapper = WorkbookWrapper()

    for workbook_basename, sheets in data.items():
        workbook_path = 'data/{}.xlsx'.format(workbook_basename)
        output_path = 'data/{}_avg.xlsx'.format(workbook_basename)
        print('Workbook: ' + workbook_path)
        wrapper.open_workbook(workbook_path)

        for sheet_name, sheet in sheets.items():
            print('Sheet: ' + sheet_name)
            wrapper.set_sheet(sheet_name)

            for data_set_n, data_map in sheet.items():
                print('Data set: ' + str(data_set_n))
                data_set = data_set_from_map(data_map)
                validate_data_set(data_set)
                wrapper.process_data_set(data_set)

        wrapper.save_workbook(output_path)


if __name__ == '__main__':
    main()
