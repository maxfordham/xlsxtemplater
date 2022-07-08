import sys
import os
import re
import getpass
import datetime
import re
import pandas as pd
FILENAME_FORBIDDEN_CHARACTERS = {"<", ">", ":", '"', "/", "\\", "|", "?", "*"}



# mf packages 
# TODO - remove this dedendency if opensource
try:
    from mf_file_utilities.applauncher_wrapper import go as _open_file
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
    if sys.platform == 'linux':
        _open_file(filename)
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
    if not isinstance(fdir, str):
        fdir = str(fdir)
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


def modify_string(s, 
                  remove_forbidden_chars=True, 
                  replace_spaces=None, 
                  fn_on_string=None,
                  min_length=None,
                  max_length=None):
    """
    
    Reference:
        [naming-a-file](https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file)
    """
    if replace_spaces is not None:
        s = s.replace(" ", replace_spaces)
    if remove_forbidden_chars:
        for c in FILENAME_FORBIDDEN_CHARACTERS:
            s = s.replace(c, "")
    if fn_on_string is not None:
        s = fn_on_string(s)
    if min_length is not None:
        if len(s) < min_length:
            s = s + "-"*(min_length-len(s))
    if max_length is not None:
        if len(s) > max_length:
            s = s[0:max_length]
    return s


