import excelbind

excelbind.register('my_function', ['x'], ['float'], 'float')
excelbind.register('my_function_2', ['s'], ['str'], 'str')

def my_function(x):
    return x*x

def my_function_2(s):
    return 'Hello ' + s
