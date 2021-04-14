

# not sure about this...

# see here:
# https://stackoverflow.com/Questions/16981921/relative-imports-in-python-3
# for discussion
#import os, sys; sys.path.append(os.path.dirname(os.path.realpath(__file__))) #  not sure if this works when executing a script from jupyter

#  add here for simplified api
from templater import to_excel
from utils import from_excel

from _version import get_versions
__version__ = get_versions()['version']
del get_versions