import sys
import os
import re
import getpass
import datetime
import re

# mf packages
try:
    from mf_file_utilities import applauncher_wrapper as aw
except:
    pass

#  extracted from mf_modules ##################################
#  from mf_modules.file_operations import open_file
def open_file(filename):
    """Open document with default application in Python."""
    if sys.platform == 'linux' and str(type(aw))== "<class 'module'>":
        aw.open_file(filename)
        #  note. this is an MF custom App for opening folders and files
        #        from a Linux file server on the local network
    else:
        try:
            os.startfile(filename)
        except AttributeError:
            subprocess.call(['open', filename])

#  from mf_modules.file_operations import jobno_fromdir

def jobno_fromdir(fdir):
    '''
    returns the job number from a given file directory

    Args:
        fdir (filepath): file-directory
    Returns:
        job associated to file-directory
    Code:
        re.findall("[J][0-9][0-9][0-9][0-9]", txt)
    '''
    matches = re.findall("[J][0-9][0-9][0-9][0-9]", fdir)
    if len(matches) == 0:
        job_no = 'J4321'
    else:
        job_no = matches[0]
    return job_no
##############################################################

def get_user():
    return getpass.getuser()

def date():
    return datetime.datetime.now().strftime('%Y%m%d')