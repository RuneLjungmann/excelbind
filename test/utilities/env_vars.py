import os
import sys
import pathlib
from contextlib import contextmanager

@contextmanager
def set_env_vars(module_name):
    try:
        backup_env = os.environ.copy()
        os.environ['EXCELBIND_MODULEDIR'] = str(pathlib.Path(__file__).parent / '..' / '..' / 'examples')
        os.environ['EXCELBIND_MODULENAME'] = module_name
        os.environ['EXCELBIND_FUNCTIONPREFIX'] = 'excelbind'
        if hasattr(sys, 'real_prefix'):
            os.environ['EXCELBIND_VIRTUALENV'] = sys.prefix

        yield
    finally:
        os.environ = backup_env
