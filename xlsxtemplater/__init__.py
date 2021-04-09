from ._version import get_versions
__version__ = get_versions()['version']
del get_versions

# not sure about this...
from xlsxtemplater import xlsxtemplater
from xlsxtemplater import templaterdefs