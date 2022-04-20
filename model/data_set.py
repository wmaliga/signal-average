import re


class DataSet:
    def __init__(self):
        self.labels = []
        self.collections_cells = {}
        self.averages_cells = {}
        self.deviations_cells = {}


def data_set_from_map(data_map):
    entries = list(data_map.keys())

    data_set = DataSet()

    for entry in entries:
        if len(entry) == 1:
            label = entry
            data_set.labels.append(label)
            data_set.collections_cells[label] = data_map[entry]
            continue

        label, typ = entry.split('_')

        if typ == 'avg':
            data_set.averages_cells[label] = data_map[entry]
        if typ == 'dev':
            data_set.deviations_cells[label] = data_map[entry]

    return data_set


def validate_data_set(data_set):
    for label, cell in data_set.collections_cells.items():
        print('{}: {}'.format(label, ','.join(cell)))

    if len(data_set.labels) < 2:
        raise ValueError('At least two collections need to be declared!')

    for label in data_set.averages_cells.keys():
        if label not in data_set.labels:
            raise ValueError('Average "%s" has no defined collection!' % label)

    for label in data_set.deviations_cells.keys():
        if label not in data_set.labels:
            raise ValueError('Deviation "%s" has no defined collection!' % label)

    collections_count = len(data_set.collections_cells[data_set.labels[0]])

    for label in data_set.labels:
        if len(data_set.collections_cells[label]) != collections_count:
            raise ValueError('Collection "%s" has different length!' % label)

    for cells in data_set.collections_cells.values():
        for cell in cells:
            __validate_cell(cell)

    for cell in data_set.averages_cells.values():
        __validate_cell(cell)

    for cell in data_set.deviations_cells.values():
        __validate_cell(cell)


def __validate_cell(cell):
    if not re.match('^[A-Z]{1,3}[0-9]{1,3}$', cell):
        raise ValueError('Cell "%s" is invalid!' % cell)
