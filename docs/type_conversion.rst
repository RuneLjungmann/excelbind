Type conversion
===============
One of the main jobs for Excelbind is to convert data types between Excel and the exposed python code.
If you provide type annotations with the python functions you expose,
then Excelbind will use that information to convert Excel types to Python types and vice versa.

If a function  is missing type annotation, then Excelbind will try to infer the type from the input.

Currently Excelbind handles floats, strings, datetimes, lists, dictionaries and ndarrays.

Auto discovery
--------------
The ``float``, ``string``, ``list`` and ``dict`` types are all auto discovered.

You should however strive to use type annotations when possible.

- It will give better performance, as Excelbind does not need to guess the type from the input, but knows it up front.
- For floats and strings Excel will do an initial type check, so if you tried to pass a string when your type annotation said float, Excel will not try to call the function, but return an error right away.

Note that the *execute_python* function (See the :doc:`examples` section) only works with types that are auto discoverable.

Type details
------------

Lists and dictionaries
^^^^^^^^^^^^^^^^^^^^^^
Excelbind will recognise a type as a `` list`` or ``dict`` if the type annotation says *List* or *Dict*.
So specifically you should not specify the element/key/value types,
instead these are assumed to be either floats or strings and a given container can hold a mix of these.

Excelbinds auto discovery rules will recognise a range as a ``list``, if either the number of rows or the number columns is one.
It will recognise a range as a ``dict``, if either the number of rows or number of columns is two.

Numpy arrays
^^^^^^^^^^^^
Excelbind will recognise a type as ``ndarray`` if the type annotation says *ndarray*, *np.ndarray* or *numpy.ndarray*.
The numpy arrays are assumed to be (dense) two-dimensional matrices.

datetime
^^^^^^^^
Excelbind will recognise a type as ``datetime`` if the type annotation says *datetime* or *datetime.datetime*.
Auto discovery is currently not supported, as Excel passes dates and datetimes as numbers to the xll layer.

Note that due to `Excels special handling of the year 1900 as a leap year
<https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year>`_,
and in order to simplify the internal implementation, Excelbind will not convert days between 1900-01-01 and 1900-03-01 correctly.
In an Excel context this is most likely not a problem.

Pandas data structures
^^^^^^^^^^^^^^^^^^^^^^
Excelbind supports both Pandas Series and DataFrames.
It will recognise a type as ``Series`` if the type annotation says *Series*, *pd.Series* or *pandas.Series*.
It will recognise a type as ``DataFrame`` if the type annotation says *DataFrame*, *pd.DataFrame* or *pandas.DataFrame*

For both data structures it is optional to supply a row/column of indices..
