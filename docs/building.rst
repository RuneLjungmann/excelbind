Building the xll Add-in yourself
================================

The project uses CMake to properly support different versions of Visual Studio.

There is a small batch file available to run cmake and build the solution.

To use this, start a visual studio command prompt (Search for 'Developer Command Prompt') and type e.g.::

    make 2022 Release x64

To create solution files and build the solution using Visual Studio 2022.

The solution file will now be located in build/2022x64 and you can open it in the Visual Studio IDE.

The excelbind.xll is now located under Binaries/x64/Release subfolder.

Running the tests
-----------------
If you haven't already done so, start by installing the pipenv tool in your python environment::

    pip install pipenv

Now create the virtual environment used for testing. Go to the top excelbind folder and type::

    pipenv install

You can now run the test by typing::

    pipenv run pytest

The test assumes that you have build the excelbind.xll in x64 release mode.
You can manually change the tests to use the debug build or one of the x86 builds, by changing the xll path in test/conftest.py.

Note however, that unless you have a debug build of numpy,
anything that imports numpy in debug mode will fail due to issues loading the underlying numpy dlls.
