import pathlib
import pytest


@pytest.fixture(scope='module')
def xll_addin_path():
    return str(pathlib.Path(__file__) / '..' / '..' / 'Release' / 'excelbind.xll')
