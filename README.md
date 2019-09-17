# Excelbind - expose your python code in Excel

Excelbind is an Excel Add-in, that uses the Excel xll api to expose your python code to Excel users in an easy and userfriendly way.

You can both expose already existing python functions in Excel as well as run python code directly from Excel.

To expose a python function in Excel as a user-defined function (UDFs) just use the excelbind function decorator:

    import excelbind
    
    @excelbind.function
    def my_simple_python_function(x: float, y: float) -> float:
        return x + y

To write python code directly inside Excel just use the execute_python function:

    =execute_python("return arg1 + arg2", A1, B1)


The project mainly piggybacks on two other projects:
 - [pybind11](https://github.com/pybind/pybind11), which provides a modern C++ api towards python
 - [xll12](https://github.com/keithalewis/xll12), which provides a modern C++ wrapper around the xll api and makes it easy to build Excel Add-ins.

Excelbind basically ties these two together to expose python to Excel users.

## Getting started

### Requirements
 - Visual Studio 2015 or higher
 - Python 3.5 or higher with the pipenv package installed 
 - Excel 2007 or higher 

See the 'limitations' section below for details on these requirements.

### Initial configuration
Out of the box, the project configuration assumes you are using Visual Studio 2019 and Python 3.7 and that python is installed in the directory C:\python37 (C:\python37_x64 for x64).

If this is not the case, you will have to make some minor adjustment to the project and property files.

If you use an older version of Visual Studio you will have to change the 'PlatformToolset' property to v140 (Visual Studio 2015) or v141 (Visual Studio 2017).

If you are using an older version of python, or python is not installed in the expected location, you need to update the path in the properties files in excelbind/src/excelbind/project_properties


### Building the xll Add-in
You can either open the solution file directly in the Visual Studio IDE and build there, or build from the command line. 

To do this, start a visual studio command prompt (i.e. run vcvars32.bat as explained in https://docs.microsoft.com/en-us/cpp/build/building-on-the-command-line?view=vs-2019)

Go to the top excelbind folder and type:

    devenv excelbind.sln /build "Release|x86" (...or "Debug|x86", "Release|x64" etc.)

The excelbind.xll is now located in the Release (or Debug) subfolder. For x64 in x64/Release (or x64/Debug).

### Running the tests
If you haven't already done so, start by installing the pipenv tool in your python environment:

    pip install pipenv

Now create the virtual environment used for testing. Go to the top excelbind folder and type: 

    pipenv install

You can now run the test by typing

    pipenv run pytest

The test assumes that you have build the excelbind.xll in x86 release mode. 
You can manually change the tests to use the debug build or one of the x64 builds, by changing the xll path in test/conftest.py.

Note however, that unless you have a debug build of numpy, 
anything that imports numpy in debug mode will fail due to issues loading the underlying numpy dlls.


### Using the Add-in from Excel
To start using Excelbind, just
- Build the excelbind.xll as explained above.
- Create a python module you want to expose - see examples/basic_functions.py for an example.
- Create a excelbind.conf config file and place it in your user home dir (pointed to by the USERPROFILE environment variable). See examples/excelbind.conf for details on the layout of the config file. Update the settings as appropriate to point to your newly made python module.
- Start Excel and add the Add-in!

## Limitations

#### Excel
The project only supports Excel 2007 or higher. There is no plan to support earlier versions of Excel (The Excel12 xll api is assumed available).

#### C++ compiler
The implementation contains a minor part which is compiler specific (assumptions on the calling conventions for calling a non-virtual method). 
This has only been tested in Visual Studio 2019, but should at least work for Visual Studio 2015 and above (and most likely also for earlier versions).

#### Python
Only support for python 3.5 and above - as type annotations are used to expose type information to Excel.

## Roadmap
See the [TODO](TODO.md) file.

## Author
Rune Ljungmann

## License
[MIT](LICENSE.txt)

## Acknowledgments
As already mentioned above, Excelbind builds heavily on the work done in 
[pybind11](https://github.com/pybind/pybind11) and 
[xll12](https://github.com/keithalewis/xll12)

A big source of inspiration is the [ExcelDNA](https://github.com/Excel-DNA/ExcelDna) project. A very mature project, which provides Excel xll bindings for C#. 
Especially the semi-dynamic exposure of functions in the dll api using a thunks table is inspired by the way this is handled in ExcelDNA.

Excelbind uses a [headers only version](https://github.com/jtilly/inih) 
of [inih](https://github.com/benhoyt/inih) for parsing ini files.
