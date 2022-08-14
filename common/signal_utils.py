import math

import numpy
from numpy import ndarray
from scipy import interpolate


def signal_average(x_values: list[ndarray], y_values: list[ndarray]):
    if len(x_values) != len(y_values):
        raise Exception('Both X and Y lists should have equal length')

    x_start = x_values[0][0]
    x_end = x_values[0][-1]
    samples = len(x_values[0])

    for i in range(len(x_values)):
        x_start = min(x_start, x_values[i][0])
        x_end = max(x_end, x_values[i][-1])
        samples = max(samples, len(x_values[i]))

    y_interpolations = []

    for i in range(len(x_values)):
        y_interpolation = interpolate.interp1d(x_values[i], y_values[i], bounds_error=False)
        y_interpolations.append(y_interpolation)

    x_avg_values = []
    y_avg_values = []
    y_dev_values = []

    for x in numpy.linspace(x_start, x_end, num=samples):
        y_interp_values = [y_interp(x) for y_interp in y_interpolations if not math.isnan(y_interp(x))]
        y_avg = sum(y_interp_values) / len(y_interp_values)
        y_dev = math.sqrt(sum([(v - y_avg) ** 2 for v in y_interp_values]) / len(y_interp_values))

        x_avg_values.append(x)
        y_avg_values.append(y_avg)
        y_dev_values.append(y_dev)

    return x_avg_values, y_avg_values, y_dev_values
