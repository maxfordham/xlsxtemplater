import sys
import os
import re
import getpass
import datetime
import re
import pandas as pd

# mf packages 
# TODO - remove this dedendency if opensource
try:
    from mf_file_utilities import applauncher_wrapper as aw
except:
    pass

def get_user():
    return getpass.getuser()

def date():
    return datetime.datetime.now().strftime('%Y%m%d')

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

def xlsxtemplated_check(fpth):
    from openpyxl import load_workbook
    wb = load_workbook(fpth)
    if wb.properties.keywords is not None and 'xlsxtemplater' in wb.properties.keywords:
        return True
    else:
        return False


def from_excel(fpth):
    """
    reads back in pandas tables that have been output using xlsxtemplater.to_excel
    Args:
        fpth(str): xl fpth
    Returns:
        li(list): of the format below
            li = {'sheet_name':'name','description':'dataframe description','df':'pd.DataFrame'}
    """
    if not xlsxtemplated_check(fpth):
        print('{} --> not created by xlsxtemplater'.format(fpth))
        return None
    cols = ['sheet_name','description']
    df_readme = pd.read_excel(fpth,sheet_name='readme')
    li = []
    for index, row in df_readme.iterrows():
        tmp = row.to_dict()
        tmp['df'] = pd.read_excel(fpth,sheet_name=row.sheet_name)
        li.append(tmp)
    return li



