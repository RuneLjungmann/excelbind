.. meta::
   :google-site-verification: 3-t0RDVvZu0ybpiusc6No1abuvIRCuKAygdTMBsJWdc

Excelbind - Expose your python code in Excel
============================================

Excelbind is a free open-source Excel Add-in, that expose your python code to Excel users in an easy and userfriendly way.

Excelbind uses the Excel xll api together with an embedded python interpreter,
so your python code runs within the Excel process. This means that the overhead of calling a python function is just a few nano seconds.

You can both expose already existing python functions in Excel as well as run python code directly from Excel.

To expose a python function in Excel as a user-defined function (UDFs) just use the excelbind function decorator::

    import excelbind

    @excelbind.function
    def my_simple_python_function(x: float, y: float) -> float:
        return x + y

To write python code directly inside Excel just use the execute_python function::

    =execute_python("return arg1 + arg2", A1, B1)

For more info on how to use Excelbind see the :doc:`getting_started` section.

The code and Add-in available at `GitHub <https://github.com/RuneLjungmann/excelbind>`_.

Requirements
------------
- Excel 2007 or higher
- Python 3.10 or higher

If you want to build the project yourself you will also need:

- Visual Studio 2022
- CMake version 3.10 or higher
- The pipenv python package installed

Acknowledgments
---------------
The project mainly piggybacks on two other projects:

- `pybind11 <https://github.com/pybind/pybind11>`_, which provides a modern C++ api towards python
- `xll12 <https://github.com/keithalewis/xll12>`_, which provides a modern C++ wrapper around the xll api and makes it easy to build Excel Add-ins.

Excelbind basically ties these two together to expose python to Excel users.

A big source of inspiration is the
`ExcelDNA <https://github.com/Excel-DNA/ExcelDna>`_ project.
A very mature project, which provides Excel xll bindings for C#.
Especially the semi-dynamic exposure of functions in the dll api
using a thunks table is inspired by the way this is handled in ExcelDNA.

Excelbind also uses:

- A `headers only version <https://github.com/jtilly/inih>`_ of `inih <https://github.com/benhoyt/inih>`_ for parsing ini files.
- The `date library <https://github.com/HowardHinnant/date>`_ by Howard Hinnant.

There are at least two other projects solving a similar problem to Excelbind:

- `xlwings <https://www.xlwings.org/>`_ using a COM-server approach rather than using xll+an embedded interpreter.
- `pyxll <https://www.pyxll.com/>`_ which is closed-source.


.. toctree::
   :hidden:
   :maxdepth: 2
   :caption: Contents:

   getting_started
   examples
   type_conversion
   building
   troubleshooting