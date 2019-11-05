# -*- coding: utf-8 -*-
"""
Created on Tue Mar 12 07:10:48 2019

@author: gbrunkhorst
"""

# takes a dataframe makes a formatted excel file 

import xlsxwriter
import numpy as np
import math
import pandas as pd
import time

# significant digit format function
def order_mag_format(x, sig_dig = 2):
    try:
        om = math.ceil(math.log(abs(x),10))-1
        if om >= 1:
            return '#,##0'
        else:
            zerostr = '0'*(sig_dig-1-om)
            return ('0.'+zerostr)
    except:
        return ('0.0')

# adjust the file name to add the date and .xlsx if not there
def adjust_file_name(file_name, add_date):
    if add_date == True:
        datestr = time.strftime("_%Y-%m-%d")
    else:
        datestr = []
    file_name = file_name.split('.xlsx')[0] + datestr + '.xlsx'
    return file_name
    
    
#format and make the excel file function
def to_excel(df, file_name = 'gbxltable_output.xlsx', 
                          sig_dig = 2, add_date = True):
    
    # adjust the file name to add the date and .xlsx if not there
    file_name = adjust_file_name(file_name, add_date)
    
    # instanciate the workbook and worksheet    
    workbook = xlsxwriter.Workbook(file_name, {'nan_inf_to_errors': True} )
    worksheet = workbook.add_worksheet()
    
    # set the header format and the string format.
    header_form = workbook.add_format({'bold': True, 'text_wrap':True, 
                                  'border':1, 'align':'center',
                                 'valign':'bottom'})
    string_form = workbook.add_format({'border':1, 'align':'center',
                                 'valign':'center'})
    
    #format the header
    for c, heading in enumerate(df.columns):
        worksheet.write(0, c, heading, header_form)
    
    #format the data rows
    for c, column in enumerate(df.columns):
        #set the minimum width and the maximum width
        width = 10
        max_width = 30
        
        for r, row in enumerate(df.index):
        
            value = df.loc[row,column]
            #strings
            if type(value) == str:
                worksheet.write((r+1), c, value, string_form)
                
                # update width as needed for strings
                cell_width = len(value)
                if cell_width > width:
                    width = min(max_width, cell_width)
                    
            #values
            elif pd.isnull(value)==False:
                num_form = workbook.add_format({'border':1, 'align':'center',
                                 'valign':'center', 
                                 'num_format':str(order_mag_format(value, sig_dig))})
                worksheet.write(r+1, c, value, num_form)
            #nulls                
            else:
                worksheet.write((r+1), c, '', string_form)
        
        worksheet.set_column(c, c, width)
    
    workbook.close()