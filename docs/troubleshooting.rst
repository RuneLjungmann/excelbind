Troubleshooting
===============

- When you open the excelbind.xll in Excel, you get the error: 'The file format and extension of 'excelbind.xll' does not match...
    This is usually because the xll can't find a dependency. Most likely you don't have your Python installation in your path, or there is a mismatch between the xll and the python installation in your path. I.e. one is x86 and the other x64.

- I don't know if I should use the x86 or x64 version of excelbind.xll
    Basically this is determined by your Excel installation.

    The way to identify if your Excel installation is x86 or x64 differs
    a bit between different Excel versions.
    In Excel 365 you go File -> Account -> About Excel and then line 2 will tell you if it is a
    32-bit installation (= x86) or a 64-bit installation (= x64).

    Note that your Python installation (the one in your PATH) needs to match the Excel installation.

    To identify what Python installation you are running, open a command prompt and type::

        python

    Then the first line will say if it is 32-bit (= x86) or 64-bit (= x64).

- You get a build error 'The project file ".../src/thirdparty/xll12/xll12.vcxproj" was not found.
    You most likely didn't pull the git submodules in src/thirdparty.

- You get a popup with an import error related to numpy when you load the excelbind.xll. The error message contains 'No module named "numpy.core._multiarray_umath"'
    This happens when you import numpy from a debug build. You will either have to skip the import while debugging, use release build or maybe try to build numpy manually in debug mode.

