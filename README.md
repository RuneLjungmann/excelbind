# Excelbind - expose your python code in Excel

[![Documentation Status](https://readthedocs.org/projects/excelbind/badge/?version=latest)](https://excelbind.readthedocs.io/en/latest/?badge=latest)
[![Build Status](https://dev.azure.com/runeljungmann/excelbind/_apis/build/status/RuneLjungmann.excelbind?branchName=master)](https://dev.azure.com/runeljungmann/excelbind/_build/latest?definitionId=1&branchName=master)

Excelbind is a free open-source Excel Add-in, that expose your python code to Excel users in an easy and userfriendly way.

Excelbind uses the Excel xll api together with an embedded python interpreter,
so your python code runs within the Excel process. This means that the overhead of calling a python function is just a few nano seconds.

You can both expose already existing python functions in Excel as well as run python code directly from Excel.

To expose a python function in Excel as a user-defined function (UDFs) just use the excelbind function decorator:

    import excelbind
    
    @excelbind.function
    def my_simple_python_function(x: float, y: float) -> float:
        return x + y

To write python code directly inside Excel just use the execute_python function:

    =execute_python("return arg1 + arg2", A1, B1)

## Documentation

The Excelbind documentation can be found at https://excelbind.readthedocs.io/en/latest/

## Roadmap
See the [TODO](TODO.md) file.

## Author
Rune Ljungmann

## License
[MIT](LICENSE.txt)

## Acknowledgments
The project mainly piggybacks on two other projects:

- [pybind11](https://github.com/pybind/pybind11), which provides a modern C++ api towards python
- [xll12](https://github.com/keithalewis/xll12), which provides a modern C++ wrapper around the xll api and makes it easy to build Excel Add-ins.

Excelbind basically ties these two together to expose python to Excel users.

A big source of inspiration is the [ExcelDNA](https://github.com/Excel-DNA/ExcelDna>) project.
A very mature project, which provides Excel xll bindings for C#.
Especially the semi-dynamic exposure of functions in the dll api
using a thunks table is inspired by the way this is handled in ExcelDNA.

Excelbind also uses:

- A [headers only version](https://github.com/jtilly/inih)
of [inih](https://github.com/benhoyt/inih) for parsing ini files
- The [date library](https://github.com/HowardHinnant/date) by Howard Hinnant.

There are at least two other projects solving a similar problem to Excelbind:

- [xlwings](https://www.xlwings.org/) using a COM-server approach rather than using xll+an embedded interpreter.
- [pyxll](https://www.pyxll.com/) which is closed-source.
