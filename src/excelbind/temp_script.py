import numpy as np
import excelbind


excelbind.register('det', ['x'], ['np.ndarray'], 'float')
excelbind.register('inv', ['x'], ['np.ndarray'], 'np.ndarray')
excelbind.register('my_function', ['x'], ['float'], 'float')
excelbind.register('my_function_2', ['s'], ['str'], 'str')
excelbind.register('my_function_3', ['x', 'y'], ['float', 'float'], 'float')
excelbind.register('my_function_4', ['s1', 's2'], ['str', 'str'], 'str')

def det(x):
    return np.linalg.det(x)

def inv(x):
    return np.linalg.inv(x)

def my_function(x):
    return x*x

def my_function_2(s):
    return 'Hello ' + s

def my_function_3(x, y):
    return x*y

def my_function_4(s1, s2):
    return 'Hello ' + s1 + s2
