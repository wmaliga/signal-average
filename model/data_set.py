LABEL_LENGTH = 1
RESULT_LENGTH = 5


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
        if len(entry) == LABEL_LENGTH:
            label = entry
            data_set.labels.append(label)
            data_set.collections_cells[label] = data_map[entry]
            continue

        label, typ = entry.split('_')

        if typ == 'avg':
            data_set.averages_cells[label] = data_map[entry]
        if typ == 'dev':
            data_set.deviations_cells[label] = data_map[entry]

    for label, cell in data_set.collections_cells.items():
        print('{}: {}'.format(label, ','.join(cell)))

    return data_set
