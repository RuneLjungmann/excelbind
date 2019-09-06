import excelbind

excelbind.register('my_function', 'float')
excelbind.register('my_function_2', 'str')

def my_function(x):
    return str(x*x)

def my_function_2(s):
    return 'Hello ' + s
