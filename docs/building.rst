Building the xll Add-in yourself
================================


Initial configuration
---------------------
Out of the box, the project configuration assumes you are using Visual Studio 2019 and Python 3.7 and that python is installed in the directory C:\python37 (C:\python37_x64 for x64).

If this is not the case, you will have to make some minor adjustment to the project and property files.

If you use an older version of Visual Studio you will have to change the 'PlatformToolset' property to v140 (Visual Studio 2015) or v141 (Visual Studio 2017).

If you are using an older version of python, or python is not installed in the expected location, you need to update the path in the properties files in excelbind/src/excelbind/project_properties

Building the xll Add-in
-----------------------
You can either open the solution file directly in the Visual Studio IDE and build there, or build from the command line.

To do this, start a visual studio command prompt (i.e. run vcvars32.bat as explained in https://docs.microsoft.com/en-us/cpp/build/building-on-the-command-line?view=vs-2019)

Go to the top excelbind folder and type::

    devenv excelbind.sln /build "Release|x86" (...or "Debug|x86", "Release|x64" etc.)

The excelbind.xll is now located in the Release (or Debug) subfolder. For x64 in x64/Release (or x64/Debug).

Running the tests
-----------------
If you haven't already done so, start by installing the pipenv tool in your python environment::

    pip install pipenv

Now create the virtual environment used for testing. Go to the top excelbind folder and type::

    pipenv install

You can now run the test by typing::

    pipenv run pytest

The test assumes that you have build the excelbind.xll in x86 release mode.
You can manually change the tests to use the debug build or one of the x64 builds, by changing the xll path in test/conftest.py.

Note however, that unless you have a debug build of numpy,
anything that imports numpy in debug mode will fail due to issues loading the underlying numpy dlls.
