from dataclasses import dataclass, asdict, field
from typing import Optional, List, Dict, Callable
import pandas as pd
import xlsxwriter as xw

#import xlsxtemplater
from xlsxtemplater.utils import get_user
from xlsxtemplater._version import get_versions

__version__ = get_versions()['version']
NAME_VERSION = 'xlsxtemplater'+'-{}'.format(__version__)

def load_colours():
    colours = {
        'ifcAqua': '#2da4a8',
        'ifcPurple': '#b72893',
        'ifcRed': '#e70051',
        'ifcBlue': '#005ca3',
        'mfSalmon': '#F7B799',
        'mfYellow': '#FFFF99',
        'nbsPurple': '#403151'
    }
    return colours

def load_formats():
    colours = load_colours()
    formats = {
        'ifcBlue': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'white',  # r64 g49 b81 #Color 21
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': colours['ifcBlue'],
            'border': 1},
        'ifcRed': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'white',  # r64 g49 b81 #Color 21
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': colours['ifcRed'],
            'border': 1},
        'ifcPurple': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': False,
            'italic': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': colours['ifcPurple'],  # #F2F2F2 = salmon,
            'border': 1},
        'ifcAqua': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': False,
            'italic': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': colours['ifcAqua'],  # yellow #F8EB2A
            'border': 1},
        'mfSalmon': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': True,
            'italic': False,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': colours['mfSalmon'],  # yellow #F8EB2A
            'border': 1},
        'border': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': False,
            'italic': False,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': 'white',
            'border': 1},
        'readme': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': False,
            'italic': False,
            'text_wrap': True,
            'valign': 'top',
            'align': 'left',
            'fg_color': 'white',
            'border': 0},
        'readme1': {
            'font_name': 'Calibri',
            'font_size': 11,
            'font_color': 'black',  # r64 g49 b81 #Color 21
            'bold': False,
            'italic': False,
            'text_wrap': True,
            'valign': 'top',
            'align': 'left',
            'fg_color': 'white',
            'border': 0}
    }
    return formats

@dataclass
class SetCol:
    """
    define column formatting

    Reference:
        https://xlsxwriter.readthedocs.io/worksheet.html
        properties taken from the worksheet.set_column() method
    """
    first_col: int
    last_col: int
    width: float
    cell_format: Dict = field(default_factory=dict)
    options: Dict = field(default_factory=dict)

@dataclass
class SetRow:
    """
    define column formatting

    Reference:
        https://xlsxwriter.readthedocs.io/worksheet.html
        properties taken from the worksheet.set_column() method
    """
    row: int
    height: float
    cell_format: Dict = field(default_factory=dict)
    options: Dict = field(default_factory=dict)

@dataclass
class Conditional:
    """
    Apply conditional formatting

    Reference:
        https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
        https://xlsxwriter.readthedocs.io/example_conditional_format.html#ex-cond-format
    """
    range: tuple = (1, 1, 1, 1)
    options: Dict = field(default_factory=dict)

@dataclass
class Textbox:
    """
    Apply conditional formatting

    Reference:
        https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
    """
    row: int
    col: int
    text: str
    options: Dict = field(default_factory=dict)

# defaults
def default_header_row_only(): #NOT IN USE
    di = {
        'row': 1,
        'height': 80,
        'cell_format': {
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            # 'fg_color': '#D7E4BC',
            'border': 1
        },
    }
    header_row = from_dict(data=di,data_class=SetRow)
    return header_row

@dataclass
class XlsxTable:
    """
    Args:
        freeze: https://xlsxwriter.readthedocs.io/example_panes.html
        table_style: https://xlsxwriter.readthedocs.io/working_with_tables.html
        col_formatting: https://xlsxwriter.readthedocs.io/worksheet.html
            worksheet.set_column() method
        row_formatting: worksheet.set_row() method
        conditional_formatting: https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
        text_box: https://xlsxwriter.readthedocs.io/example_textbox.html#ex-textbox
    """
    freeze: tuple = (1,1)
    table_style: str = 'Table Style Light 8'
    col_formatting: List[SetCol] = field(default_factory=list)
    row_formatting: List[SetRow] = field(default_factory=list)# field(default_factory=lambda: [default_header_row_only()])
    conditional_formatting: List[Conditional] = field(default_factory=list)
    text_box: List[Textbox] = field(default_factory=list)
    hide_grid: bool = True

    def __post_init__(self):
        #  apply default header row formatting
        self.row_formatting.insert(0, SetRow(0,50,{'bold': True,'text_wrap': True, 'valign': 'top','border': 1}))

@dataclass
class FileProperties:
    title: str = None
    subject: str = None
    author: str = get_user()
    manager: str = None
    company: str = 'Max Fordham'
    category: str = None
    keywords: str = ''
    comments: str = None
    status: str = None

    def __post_init__(self):
        self.keywords = self.keywords + ', ' + NAME_VERSION


def df_to_sheet_table(df: pd.DataFrame,
                      writer: pd.ExcelWriter,
                      workbook, 
                      sheet_name,
                      xlsx_params: XlsxTable = XlsxTable(),
                      inject_header_params=None
                      ):
    '''
    an xlsxwriter template for writing a pd.DataFrame to an excel sheet
    as a table with customisable formatting.
    this is the backbone sheet template that is used by all the other ones.
    THE INDEX IS IGNORED AND WILL NOT APPEAR IN THE OUTPUT.

    Args:
        df (pd.DataFrame):
        writer (class): xlsxwriter object
        sheet_name (string):
        xlsx_params (XlsxTable); default= XlsxTable(): defines formatting of the output excel file
        inject_header_params (dict): default=None, can be used to inject in header vars into the excel
            table. the dict keys and the column headings and the dict value is a dict to update the header
            item with. refer to: https://xlsxwriter.readthedocs.io/working_with_tables.html

            ```relevant code
            header = [{'header': col} for col in df.columns.tolist()]
            if inject_header_params is not None:
                for h in header:
                    col = h['header']
                    if col in inject_header_params.keys():
                        h.update(inject_header_params[col])
            ```

    Returns:
        worksheet (class): xlsxwriter object that defines an excel worksheet output.
            another can then compile many sheets in a single output

    Code:
        df.to_excel(writer, sheet_name, index=None)
        worksheet = writer.sheets[sheet_name]

        if xlsx_params.hide_grid == True:
            worksheet.hide_gridlines(2)

        # get table range
        end_row = len(df.index)
        last_column = len(df.columns)-1
        cell_range = xw.utility.xl_range(0, 0, end_row, last_column)

        # create table
        #df = df.reset_index()  # FIX
        header = [{'header': col} for col in df.columns.tolist()]
        if inject_header_params is not None:
            for h in header:
                col = h['header']
                if col in inject_header_params.keys():
                    h.update(inject_header_params[col])

        worksheet.add_table(cell_range,
                            {'style': xlsx_params.table_style,
                            'header_row': True,
                            'first_column': True,
                            'columns': header})

        # set col formatting
        #  worksheet.set_column(first_col, last_col, width, cell_format, options)
        for col in xlsx_params.col_formatting:
            cell_format = workbook.add_format(col.cell_format)
            worksheet.set_column(col.first_col, col.last_col, col.width, cell_format, col.options)

        # set row formatting
        for row in xlsx_params.row_formatting:
            cell_format = workbook.add_format(row.cell_format)
            worksheet.set_row(row.row, row.height, cell_format, row.options)

        # insert textboxes
        for t in xlsx_params.text_box:
            worksheet.insert_textbox(t.row, t.col, t.text, t.options)

        # insert conditional formatting
        for c in xlsx_params.conditional_formatting:
            worksheet.conditional_format(c.range, c.options)

        # freeze header row and index
        if xlsx_params.freeze is not None:
            worksheet.freeze_panes(xlsx_params.freeze[0], xlsx_params.freeze[1])

        return worksheet
    '''

    df.to_excel(writer, sheet_name, index=None)
    worksheet = writer.sheets[sheet_name]

    if xlsx_params.hide_grid == True:
        worksheet.hide_gridlines(2)

    # get table range
    end_row = len(df.index)
    last_column = len(df.columns)-1
    cell_range = xw.utility.xl_range(0, 0, end_row, last_column)

    # create table
    #df = df.reset_index()  # TODO
    header = [{'header': col} for col in df.columns.tolist()]
    if inject_header_params is not None:
        for h in header:
            col = h['header']
            if col in inject_header_params.keys():
                h.update(inject_header_params[col])

    worksheet.add_table(cell_range,
                        {'style': xlsx_params.table_style,
                         'header_row': True,
                         'first_column': True,
                         'columns': header})

    # set col formatting
    #  worksheet.set_column(first_col, last_col, width, cell_format, options)
    for col in xlsx_params.col_formatting:
        cell_format = workbook.add_format(col.cell_format)
        worksheet.set_column(col.first_col, col.last_col, col.width, cell_format, col.options)

    # set row formatting
    for row in xlsx_params.row_formatting:
        cell_format = workbook.add_format(row.cell_format)
        worksheet.set_row(row.row, row.height, cell_format, row.options)

    # insert textboxes
    for t in xlsx_params.text_box:
        worksheet.insert_textbox(t.row, t.col, t.text, t.options)

    # insert conditional formatting
    for c in xlsx_params.conditional_formatting:
        worksheet.conditional_format(c.range, c.options)

    # freeze header row and index
    if xlsx_params.freeze is not None:
        worksheet.freeze_panes(xlsx_params.freeze[0], xlsx_params.freeze[1])

    return worksheet

@dataclass
class TableObj:
    df: pd.DataFrame
    sheet_name: str = 'sheet_name'
    description: str = 'short description of the table. add more details to notes.'
    notes: Dict = field(default_factory=dict)

@dataclass
class SheetObj(TableObj):
    xlsx_params: XlsxTable = XlsxTable()
    xlsx_exporter: Callable = df_to_sheet_table


@dataclass
class ToExcel:
    sheets: List[SheetObj]


# custom definitions
def params_readme(df):
    """
    defines the parameters for the readme sheet
    """
    freeze = (1, 1)
    table_style = 'Table Style Light 8'
    colours = load_colours()
    formats = load_formats()
    end = len(list(df))
    col_formatting = [
        {
            'first_col': 0,
            'last_col': 0,
            'width': 20,
            'cell_format': formats['readme'],
            'options': {'text_wrap': True, 'align': 'left'}
        },
        {
            'first_col': 1,
            'last_col': end,
            'width': 40,
            'cell_format': formats['readme1'],
            'options': {'text_wrap': True, 'align': 'left'}
        }
    ]

    tbox = [
        {
            'row': len(df) + 2,
            'col': 1,
            'text': '''
                    This readme tab contains the meta data for the datatables in the other worksheets. \n
                    Each column corresponds to 1no sheet in this excel workbook.
                    ''',
            'options': {
                'fill': {'color': colours['mfYellow']},
                'width': 800,
                'height': 160,
                'font': {'bold': True},
            }
        }
    ]

    xlsx_params = {
        'freeze': freeze,
        'col_formatting': col_formatting,
        'textbox': tbox,
        'table_style': table_style,
        'hide_grid': True,
    }
    ParamsReadme = from_dict(data_class=XlsxTable,data=xlsx_params)
    return ParamsReadme


def params_ifctemplate():
    '''
    formatting data for bsdd ifc product data template (generated by the IfcDataTemplater app)
    '''
    table_style = 'Table Style Light 2'
    freeze = (1, 3)
    tbox = [
        {
            'row': 1,
            'col': 18,
            'text': '''
                    This is an IFC Building Data Dictionary Product Data Sheet. \n
                    Max Fordham have transposed the columns to rows to make it easier to read. \n
                    Any column with the prefix "Mf" will be ignored. \n
                    Use to Mf examples as a guide, it indicates what information should be filled in for a given product. \n
                    Over time, Mf examples will be completed for all of the products that we take responsibility for specifying. \n
                    \n
                    An "Image" row has also been added. \n
                    Copy and paste an image into this row for your product. \n
                    This Image will then be mapped to a family in your Revit model. \n
                    The "ModelReference" must equal the "TypeMark" in the Revit model.
                    ''',
            'options': {
                'fill': {'color': 'yellow'},
                'width': 800,
                'height': 160,
                'font': {'bold': True},
            }
        }
    ]
    formats = load_formats()
    col_formatting = [
        {
            'first_col': 1,
            'last_col': 1,
            'width': 40,
            'cell_format': formats['ifcBlue'],
            'options': {},
        },
        {
            'first_col': 2,
            'last_col': 2,
            'width': 60,
            'cell_format': formats['ifcBlue'],
            'options': {},
        },
        {
            'first_col': 3,
            'last_col': 3,
            'width': 8,
            'cell_format': formats['ifcRed'],
            'options': {},
        },
        {
            'first_col': 4,
            'last_col': 8,
            'width': 30,
            'cell_format': formats['ifcRed'],
            'options': {'level': 1, 'hidden': True},
        },
        {
            'first_col': 9,
            'last_col': 9,
            'width': 8,
            'cell_format': formats['ifcPurple'],
            'options': {},
        },
        {
            'first_col': 10,
            'last_col': 12,
            'width': 30,
            'cell_format': formats['ifcPurple'],
            'options': {'level': 1, 'hidden': True},
        },
        {
            'first_col': 13,
            'last_col': 13,
            'width': 8,
            'cell_format': formats['ifcAqua'],
            'options': {},
        },
        {
            'first_col': 14,
            'last_col': 16,
            'width': 30,
            'cell_format': formats['ifcAqua'],
            'options': {'level': 1, 'hidden': True},
        },
    ]
    xlsx_params = {
        'freeze': freeze,
        'col_formatting': col_formatting,
        'textbox': tbox,
        'table_style': table_style,
        'hide_grid': True,
    }
    ParamsIfc = from_dict(data_class=XlsxTable, data=xlsx_params)
    return ParamsIfc

