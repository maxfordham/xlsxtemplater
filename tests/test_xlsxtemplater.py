#!/usr/bin/env python

"""Tests for `xlsxtemplater` package."""


# import unittest
from xlsxtemplater import from_excel, to_excel
from xlsxtemplater.templater import df_to_sheet_table, params_ifctemplate
import pandas as pd
import pathlib

fpth_in = pathlib.Path('test_data/bsDataDictionary_Psets.xlsx')
fpth_out = fpth_in.parent / 'bsDataDictionary_Psets-out.xlsx'
fpth_out.unlink(missing_ok=True)
df = pd.read_excel(fpth_in)


class TestXlsxTemplater:

    def test_to_excel(self):

        di = {
            'sheet_name': 'IfcProductDataTemplate',
            'xlsx_exporter': df_to_sheet_table,
            'xlsx_params': params_ifctemplate(),
            'df': df,
        }
        li = [di]
        
        to_excel(li, fpth_out, openfile=False)
        assert fpth_out.is_file()

    # TODO: fix this! 
    def test_from_excel(self):
        li = from_excel(fpth_in)
        assert li is not None
        assert list(li[0].keys()) ==  ["sheet_name",
            "description",
            "JobNo",
            "Date",
            "Author"]


