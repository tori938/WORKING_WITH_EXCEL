#standard libraries
import pandas as pd, numpy as np

#excel helper
import xlrd

#path directory
import os

#import the user functions
from names_and_values import convert_address
from names_and_values import convert_month_to_number
from names_and_values import get_accrual_names
from names_and_values import get_accrual_values

from names_and_values import create_report_month
from names_and_values import create_report_address

from names_and_values import get_cost_names
from names_and_values import get_cost_values

from names_and_values import create_content

#store results
from collections import OrderedDict

#logging
from loguru import logger

#confidential
from dotenv import dotenv_values





#set the credentials
credentials = dotenv_values()

#create the logging file
logger.add('log/parking_n_storage_untransmitted.info',
           format="{time} {level} {message}")


#set the folder path
reports_folder_path = credentials['reports_folder_path']

all_report_files = os.listdir(reports_folder_path)
logger.info(f'current report status: {all_report_files}')


set_the_month = credentials['month']
set_the_year = credentials['year']


## DEDUCTIONS
#create an ordered dict to store deductions
deduction = OrderedDict([
    ('Дата', []),
    ('Адрес', []),
    ('Атрибут', []),
    ('Исключения', [])
    ]
)


### DEDUCTIONS
#set the folder path
deductions_folder_path = credentials['deductions_folder_path']
all_deduction_files = os.listdir(deductions_folder_path)
logger.info(f'current deduction files: {all_deduction_files}')


##DEDUCTIONS FOR BUILDING 26
for name in os.listdir(deductions_folder_path):
    #check that the relevant file is in the directory
    if name.endswith('.xls') and ('26' in name) and set_the_month in name and set_the_year in name:
        filename = os.path.join(deductions_folder_path, name)
        
        #load the workbook
        wb = xlrd.open_workbook(filename)
        logger.info(f'deduction file for building 26: {filename}')

        #store the first sheet into a variable
        sheet = wb.sheet_by_index(0)


        #find the number of rows and columns
        records = sheet.nrows
        features = sheet.ncols
        logger.info(f'number of rows: {records}, number of columns: {features}')


        #store the report cell value
        first_cell = sheet.cell_value(0, 0)
        logger.info(f'1st cell value: {first_cell}')

        #store the address value
        address_cell = sheet.cell_value(2, 0)
        logger.info(f'address: {address_cell}')

        #convert the address
        address = convert_address(address_cell)
        
        #remove paragraph breaks
        clean_cell = first_cell.replace("\n", "")

        #find the last instance of month and year in a string
        last_instance = clean_cell.split("-", 1)[1]

        #trim the beginning of the string from any spaces
        month_year_cell = last_instance.lstrip()


        #trim the string to find the year of the report
        year = int(month_year_cell[(len(month_year_cell)-4):len(month_year_cell)])

        #trim the string to find the year of the report
        month_cell = month_year_cell[:(len(month_year_cell)-5)]

        #convert month to number
        month = convert_month_to_number(month_cell)
        logger.info(f'year and month: {year}-{month}')

        #get the cost names for building 26
        costs = get_cost_names(sheet, 2, 5, features)
        logger.info(f'building 26 deductions > before removal of unnecessary elements: {costs}')

        #get the cost values for building 26
        cost_values = get_cost_values(sheet, records, 5, features)
        logger.info(f'building 26 deduction values > before removal of unnecessary elements: {cost_values}')

        #find the index of the unnecessary elements
        var1 = 'Всего'
        find_the_index = []

        for i, e in enumerate(costs):
            if e.startswith(var1):
                find_the_index.append(i)

        #remove unnecessary elements from cost names for building 26
        for e in costs:
            if e.startswith(var1):
                costs.remove(e)

        logger.info(f'building 26 deductions: {costs}')
        

        #remove unecessary elements from cost values for building 26
        for m in find_the_index:
            del cost_values[m]

        logger.info(f'building 26 deduction values: {cost_values}')

        #set the report month (and fill in)
        report_month_lst = create_report_month(month, year, cost_values)
        logger.info(f'report month: {report_month_lst[0]}')

        #set the report address (and fill in)
        address_lst = create_report_address(cost_values, address)
        logger.info(f'report address: {address_lst[0]}')


        #add results to the ordered dict       
        for d in report_month_lst:
            deduction['Дата'].append(d)
        
        for a in address_lst:
            deduction['Адрес'].append(a)

        for c in costs:
            deduction['Атрибут'].append(c)

        for v in cost_values:
            deduction['Исключения'].append(v)


### DEDUCTIONS FOR BUILDING 30
for name in os.listdir(deductions_folder_path):
    #check that the relevant file is in the directory
    if name.endswith('.xls') and ('30' in name) and set_the_month in name and set_the_year in name:
        filename = os.path.join(deductions_folder_path, name)
        
        #load the workbook
        wb = xlrd.open_workbook(filename)
        logger.info(f'deduction file for building 30: {filename}')

        #store the first sheet into a variable
        sheet = wb.sheet_by_index(0)


        #find the number of rows and columns
        records = sheet.nrows
        features = sheet.ncols
        logger.info(f'number of rows: {records}, number of columns: {features}')

        #store the report cell value
        first_cell = sheet.cell_value(0, 0)
        logger.info(f'1st cell value: {first_cell}')

        #store the address value
        address_cell = sheet.cell_value(2, 0)

        #convert the address
        address = convert_address(address_cell)
        logger.info(f'address: {address_cell}')

        #remove paragraph breaks
        clean_cell = first_cell.replace("\n", "")

        #find the last instance of month and year in a string
        last_instance = clean_cell.split("-", 1)[1]

        #trim the beginning of the string from any spaces
        month_year_cell = last_instance.lstrip()


        #trim the string to find the year of the report
        year = int(month_year_cell[(len(month_year_cell)-4):len(month_year_cell)])

        #trim the string to find the year of the report
        month_cell = month_year_cell[:(len(month_year_cell)-5)]

        #convert month to number
        month = convert_month_to_number(month_cell)
        logger.info(f'year and month: {year}-{month}')

        #get the cost names for building 30
        costs = get_cost_names(sheet, 2, 5, features)
        logger.info(f'building 30 deductions > before removal of unnecessary elements: {costs}')

        #get the cost values for building 26
        cost_values = get_cost_values(sheet, records, 5, features)
        logger.info(f'building 30 deduction values > before removal of unnecessary elements: {cost_values}')

        #find the index of the unnecessary elements
        var1 = 'Всего'
        find_the_index = []

        for i, e in enumerate(costs):
            if e.startswith(var1):
                find_the_index.append(i)

        #remove unnecessary elements from cost names for building 30
        for e in costs:
            if e.startswith(var1):
                costs.remove(e)

        logger.info(f'building 30 deductions: {costs}')
        

        #remove unecessary elements from cost values for building 30
        for m in find_the_index:
            del cost_values[m]

        logger.info(f'building 30 deduction values: {cost_values}')


        #set the report month (and fill in)
        report_month_lst = create_report_month(month, year, cost_values)
        logger.info(f'report month: {report_month_lst[0]}')

        #set the report address (and fill in)
        address_lst = create_report_address(cost_values, address)
        logger.info(f'report address: {address_lst[0]}')

        #add results to the ordered dict       
        for d in report_month_lst:
            deduction['Дата'].append(d)
        
        for a in address_lst:
            deduction['Адрес'].append(a)

        for c in costs:
            deduction['Атрибут'].append(c)

        for v in cost_values:
            deduction['Исключения'].append(v)


#write the deduction results to a dataframe
report_deductions_summary = pd.DataFrame(deduction)
logger.success(f'deduction report for {set_the_month}-{set_the_year} was created')


## MAPPING
#set the folder path
mapping_folder_path = credentials['export_folder_path']
mapping_file_name = 'Выручка по Паркингам и кладовкам - Непереданные.xlsx'

#import the table
address_table = pd.read_excel(mapping_folder_path + mapping_file_name,
                              sheet_name='address_map')

#import the table
nomenclature_table = pd.read_excel(mapping_folder_path + mapping_file_name,
                                   sheet_name='nomenclature_map')

#trim the variables (in case there are spaces in the data)
nomenclature_obj = nomenclature_table[['Услуга в отчете', 'Номенклатура 1С', 'Код номенклатуры', 'Подразделение', 'Код подразделения']]
nomenclature_table[nomenclature_obj.columns] = nomenclature_obj.apply(lambda x: x.str.lstrip().str.rstrip())


## MERGE TABLES
#merge the accrual report and the address map
deductions_address = report_deductions_summary.merge(address_table,
                                                     how='left',
                                                     left_on='Адрес',
                                                     right_on='Адрес в отчете'
                                                     )

#merge the accrual report and the address map with nomenclature map
deductions_address_nomenclature = deductions_address.merge(nomenclature_table,
                                                           how='left',
                                                           left_on='Атрибут',
                                                           right_on='Услуга в отчете'
                                                           )

#check for duplicate values
dupl_columns = list(deductions_address_nomenclature.columns)

mask = deductions_address_nomenclature.duplicated(subset=dupl_columns)
duplicates = deductions_address_nomenclature[mask]
logger.info(f'Number of Duplicates in Merged Tables: {duplicates.shape[0]}')

#drop duplicates (if there are any)
deductions_address_nomenclature = deductions_address_nomenclature.drop_duplicates(subset=dupl_columns)


#fill in deductions with 0
deductions_address_nomenclature['Исключения'] = deductions_address_nomenclature['Исключения'].fillna(0)

#rename the columns
deductions_address_nomenclature = deductions_address_nomenclature.rename(columns={'Дата': 'Период',
                                                                                  'Исключения': 'Сумма Исключений'
                                                                                  })


#apply the function
deductions_address_nomenclature['Содержание'] = deductions_address_nomenclature.apply(create_content,
                                                                                      axis=1)

#add a column
deductions_address_nomenclature['Кол-во'] = 1

deductions_address_nomenclature['Код адреса'] = deductions_address_nomenclature['Код адреса'].apply(lambda x: '000000' + str(x))


## RE-ORDER COLUMNS
#find the index of each column
idx = [deductions_address_nomenclature.columns.get_loc(c) for c in deductions_address_nomenclature.columns if c in deductions_address_nomenclature]

#re-order columns using iloc
deductions_reordered = deductions_address_nomenclature.iloc[:, [0, 7, 8, 9, 10, 11, 12, 5, 6, 14, 15, 19, 16, 17, 18, 20, 3]]


#extract the columns
clms = deductions_reordered.columns.to_list()

#excluding the last column
clms = clms[:-1]

#group by all > to merge the amounts by nomenclature
deductions_reordered = deductions_reordered.groupby(clms)['Сумма Исключений'].sum().to_frame()

#create the index > to fill down the values (repeats / duplicates)
deductions_reordered = deductions_reordered.reset_index()


def update_current_data_for_untransmitted():

    '''
        > export the current monthly results, re-write the existing results
    '''
    
    xlsx_name = 'Выручка по Паркингам и кладовкам - Непереданные.xlsx'
    sheet = 'для_загрузки'

    export_folder_path = credentials['export_folder_path']

    with pd.ExcelWriter(export_folder_path + "/" + xlsx_name,
                        engine='openpyxl',
                        mode='a',
                        if_sheet_exists='replace') as writer:
            deductions_reordered.to_excel(writer,
                                          sheet_name=sheet,
                                          index=False,
                                          header=True)

    logger.success(f'new untransmitted report for {set_the_month}-{set_the_year} is generated')


def create_monthly_data_for_untransmitted():

    '''
        > export / archive the current monthly results
    '''
    
    xlsx_name = 'Выручка по Паркингам и кладовкам - Непереданные - Архив.xlsx'
    sheet = f'для_загрузки_{set_the_month}-{set_the_year}'

    export_folder_path = credentials['export_folder_path']

    with pd.ExcelWriter(export_folder_path + "/" + xlsx_name,
                        engine='openpyxl',
                        mode='a',
                        if_sheet_exists='replace') as writer:
            deductions_reordered.to_excel(writer,
                                          sheet_name=sheet,
                                          index=False,
                                          header=True)

    logger.success(f'new historical untransmitted report for {set_the_month}-{set_the_year} is generated')

update_current_data_for_untransmitted()
create_monthly_data_for_untransmitted()