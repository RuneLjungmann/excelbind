import pathlib
import pytest
import struct


def interpreter_platform():
    return 'x64' if struct.calcsize('P') == 8 else 'Win32'


@pytest.fixture(scope='module')
def xll_addin_path():
    return str(pathlib.Path(__file__).parent / '..' / 'Binaries' / interpreter_platform() / 'Release' / 'excelbind.xll')
