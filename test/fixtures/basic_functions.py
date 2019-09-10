import numpy as np
import excelbind


excelbind.register('add', ['x', 'y'], ['float', 'float'], 'float')
excelbind.register('mult', ['x', 'y'], ['float', 'float'], 'float')
excelbind.register('det', ['x'], ['np.ndarray'], 'float')
excelbind.register('inv', ['x'], ['np.ndarray'], 'np.ndarray')
excelbind.register('concat', ['s1', 's2'], ['str', 'str'], 'str')


def add(x, y):
    return x + y


def mult(x, y):
    return x*y


def det(x):
    return np.linalg.det(x)


def inv(x):
    return np.linalg.inv(x)


def concat(s1, s2):
    return s1 + s2
