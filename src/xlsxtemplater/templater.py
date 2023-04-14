import pandas as pd
import os
import copy
import subprocess
from dataclasses import asdict
from xlsxtemplater.utils import open_file, jobno_fromdir, get_user, date, modify_string
from xlsxtemplater.templaterdefs import *


def create_meta(fpth):
    di = {}
    di['JobNo'] = jobno_fromdir(fpth)
    di['Date'] = date()
    di['Author'] = get_user()
    return di

def create_readme(toexcel: ToExcel) -> SheetObj:
    """
    creates a readme dataframe from the metadata in the dataobject definitions
    """
    def notes_from_sheet(sheet: SheetObj):
        di = {
            'sheet_name':sheet.sheet_name,
            'xlsx_params':str(type(sheet.xlsx_params)),
            'xlsx_exporter': str(sheet.xlsx_exporter)
        }
        di.update(sheet.notes)
        return di

    li = [notes_from_sheet(sheet) for sheet in toexcel.sheets]
    df = pd.DataFrame.from_records(li).set_index('sheet_name')
    df = df.reset_index()
    di = {
        'sheet_name': 'readme',
        'xlsx_exporter': df_to_sheet_table,
        'xlsx_params': params_readme(df),
        'df': df,
    }
    readme = from_dict(data_class=SheetObj,data=di)
    return readme

def create_sheet_objs(data_object, fpth, validate_sheet_name=True) -> ToExcel:
    '''
    pass a dataobject and return a ToExcel objects
    this function interprests the user input and tidies into the correct format.
    '''

    def default(df, counter):
        di_tmp = {
            'sheet_name': 'Sheet{0}'.format(counter),
            #'xlsx_exporter': df_to_sheet_table,
            #'xlsx_params': None,
            'df': df,
        }
        counter += 1
        return di_tmp, counter

    def add_notes(di, fpth):
        if 'notes' not in di.keys():
            di['notes'] = {}
        di['notes'].update(create_meta(fpth))
        return di

    counter = 1
    lidi = []
    if type(data_object) == pd.DataFrame:
        # then export the DataFrame with the default exporter (i.e. as a table to sheet_name = Sheet1)
        di, counter = default(data_object, counter)
        di = add_notes(di)
        lidi.append(di)
    if type(data_object) == list:
        # then iterate through the list. 1no sheet / item in list
        for l in data_object:
            if type(l) == pd.DataFrame:
                # then export the DataFrame with the default exporter (i.e. as a table to sheet_name = Sheet#)
                di, counter = default(l, counter)
                di = add_notes(di, fpth)
                lidi.append(di)
            elif type(l) == dict:
                # then export the DataFrame with the exporter defined by the dict
                l = add_notes(l, fpth)
                #l = add_defaults(l)
                lidi.append(l)
            else:
                print('you need to pass a list of dataframes or dicts for this function to work')
    if type(data_object) == dict:
        data_object = add_notes(data_object, fpth)
        lidi.append(data_object)
    
    if validate_sheet_name is not None:
        for l in lidi: 
            l['sheet_name'] = modify_string(l['sheet_name'], max_length=30)
        
    sheets = [from_dict(data_class=SheetObj,data=l) for l in lidi] #  defaults are added here if not previously specified
    toexcel = ToExcel(sheets=sheets)
    return toexcel

def object_to_excel(toexcel: ToExcel, fpth: str, file_properties: FileProperties):
    """
    Args:
        toexcel: ToExcel object
        fpth:
        file_properties: FileProperties object

    Returns:
        fpth
    """
    # initiate xlsxwriter
    writer = pd.ExcelWriter(fpth, engine='xlsxwriter')
    workbook = writer.book
    for sheet in toexcel.sheets:
        sheet.xlsx_exporter(sheet.df, writer, workbook, sheet.sheet_name, sheet.xlsx_params)
    workbook.set_properties(asdict(file_properties))
    # save and close the workbook
    writer.close()
    return fpth

def to_excel(data_object,
             fpth,
             file_properties: FileProperties=None,
             openfile: bool=False,
             make_readme: bool=True) -> str:
    """
    function to output dataobject (list of dicts of dataframes and associated metadata)
    to excel in nicely formatted tables. 

    Args:
        data_object (list of dicts): gets converted to a templaterdefs.ToExcel object, which is a list
            of templaterdefs.SheetObj's. any dict keys not in SheetObj definition will be ignored. 
            min required is [{'df':df}]
        fpth (str filepath): of xlsx output
        file_properties: FileProperties obj defining metadata
        openfile: bool
        make_readme: creates a readme header sheet. default to true. avoid changing unless
            necessary as it is required for the from_excel comm and. 
    
    Returns:
        fpth: of output excel file

    Example:
        #  a vanilla example
        df = pd.DataFrame.from_dict({'col1':[0,1],'col2':[1,2]})
        li = [{
            'sheet_name': 'df',
            #'xlsx_exporter': sheet_table, # don't pass xlsx_exporter to get default
            #'xlsx_params': params_ifctemplate(), # don't pass xlsx_params to get default
            'df': df,
            'notes':{
                'a note': 'a note',
                'how many notes?':'as many as you like',
                'what types?':'numbers and strings only',
                'how are they shown?':'as fields in the readme sheet'
            }
        }]
        to_excel(li, fpth, openfile=True,make_readme=True)
    """
    toexcel = create_sheet_objs(data_object, fpth)
    if make_readme:
        readme = create_readme(toexcel) # get sheet meta data
        # create metadata to make the readme worksheet
        toexcel.sheets.insert(0, readme)
    if file_properties is None:
        file_properties = FileProperties()
    object_to_excel(toexcel, fpth, file_properties)
    if openfile:
        open_file(fpth)
    return fpth

#  TODO: make to_json function that outputs the same data to a json file.

if __name__ == '__main__':
    if __debug__ == True:
        fdir = os.path.join('test_data')
        fpth = os.path.join(fdir,'bsDataDictionary_Psets.xlsx')
        df = pd.read_excel(fpth)
        #fpth = wdir + '\\' + 'bsDataDictionary_Psets-processed.xlsx'
        #df1 = pd.read_excel(fpth,sheet_name='1_PropertySets')
        di = {
            'sheet_name': 'IfcProductDataTemplate',
            'xlsx_exporter': df_to_sheet_table,
            'xlsx_params': params_ifctemplate(),
            'df': df,
        }
        li = [di]
        fpth = os.path.join(fdir,'bsDataDictionary_Psets-out.xlsx') 
        to_excel(li, fpth, openfile=False)
        print('{} --> written to excel'.format(fpth))
        from utils import from_excel
        li = from_excel(fpth)
        if type(li) is not None:
            print('{} --> read from excel'.format(fpth))
