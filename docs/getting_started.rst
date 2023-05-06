Getting Started
===============

Before you get started, you need Excel and Python installed on your machine.

- You need Excel 2007 or higher and Python 3.10 or higher.
- Make sure they are either both x86 or x64 - see the :doc:`troubleshooting` section if in doubt.
- You need the location of Python in your PATH environment variable

You are now ready to get up and running in a few steps:

- Download the Excelbind repository from `Github <https://github.com/RuneLjungmann/excelbind>`_
    Choose the green 'Clone or download' button on the right hand side to download the code.

- You need to set up a config file in your user directory.
    You can just copy the example config file examples/excelbind.conf to get started.
    The example file contains explanations of the various settings.

    Your user directory will most likely be something like 'C:/users/<your user name>/'


- You need some python functions that you want to run from Excel.
    As a first exmaple you can just pick the example file examples/excelbind_export.py
    and copy that to your user directory along with the config file above.

    You can later change the location and name of the python file you want
    to export by updating the corresponding settings in the config file.


- Open Excel and open the excelbind.xll file.
    The file is either::

        Binaries/x86/excelbind.xll or Binaries/x64/excelbind.xll

    Depending on your Excel and Python installations being x86 or x64.

    If you are in doubt about what to choose look in the :doc:`troubleshooting` section.

    If you want to build the excelbind.xll file from scratch, have a look in the :doc:`building` section.

You can now call one of the python functions from Excel like::

    =excelbind.add(2, 2)

