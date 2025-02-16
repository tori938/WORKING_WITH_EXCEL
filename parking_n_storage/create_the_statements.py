#standard libraries
import pandas as pd, numpy as np

#excel helper
import xlrd

#path directory
import os

#import the user functions
from names_and_values import convert_month_to_number
from names_and_values import get_accrual_names
from names_and_values import get_accrual_values

from names_and_values import create_report_month
from names_and_values import create_report_address

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
logger.add('log/parking_n_storage_statements.info',
           format="{time} {level} {message}")



## REPORTS
#create an ordered dict to store reports
report = OrderedDict([
    ('Дата', []),
    ('Адрес', []),
    ('Атрибут', []),
    ('Текущие Начисления', [])
    ]
)


#set the folder path
reports_folder_path = credentials['reports_folder_path']
all_report_files = os.listdir(reports_folder_path)
logger.info(f'current report status: {all_report_files}')


set_the_month = credentials['month']
set_the_year = credentials['year']


#iterate through all the files in the directory
for name in os.listdir(reports_folder_path):
    #check that the relevant file is in the directory
    if name.endswith('.xls') and set_the_month in name and set_the_year in name:
        filename = os.path.join(reports_folder_path, name)
        
        #load the workbook
        wb = xlrd.open_workbook(filename)
        logger.info(f'report file: {filename}')
    

        #store the first sheet into a variable
        sheet = wb.sheet_by_index(0)


        #find the number of rows and columns
        records = sheet.nrows
        features = sheet.ncols
        logger.info(f'number of rows: {records}, number of columns: {features}')

        #store the report cell value
        first_cell = sheet.cell_value(0, 0)
        logger.info(f'1st cell value: {first_cell}')


        #find the cells corresponding to each reporting address > store the address value
        #> an assumption that the address line exists in the first 30 rows and 1st column
        ##store the row number for address 30 to find the list of services at a later stage
        store_row = 0
        for row in range(1, 30):
            for column in range(1):
                cell = sheet.cell_value(row, column)
                if 'Суздальское д. 26 к.1' in cell:
                    address_cell26 = sheet.cell_value(row, column)
                elif 'Суздальское д. 30 к.2 стр.1' in cell:
                    address_cell30 = sheet.cell_value(row, column)
                    store_row = row
                    continue
        logger.info(f'1st address: {address_cell26}, 2nd address: {address_cell30}')
        logger.info(f'store row for address line 30: {store_row}')

        #find the substring that returns text + month and year
        misc_substring = first_cell[first_cell.find("за") + first_cell[first_cell.find("за") : len(first_cell)].find("2024") - 13 : \
                        first_cell.find("за") + first_cell[first_cell.find("за") : len(first_cell)].find("2024") + 4]
        logger.info(f'text string: {misc_substring}')

        #return the month (using the first space as the delimiter)
        month_cell = misc_substring.split(" ", 1)[1].rsplit(" ", 1)[0]

        #trim the string to find the year of the report
        year = misc_substring[-4:]

        #convert month abbr to numerical
        month = convert_month_to_number(month_cell)
        logger.info(f'year and month: {year}-{month}')

        ###REPORTS FOR BUILDING 26
        #get accruals for building 26
        accrual_names26 = get_accrual_names(sheet, 5, 19, 1)
        logger.info(f'building 26 services {accrual_names26}')
        
        #get accrual values for building 26
        current_accrual_values26 = get_accrual_values(sheet, 5, 19, 5, 6)
        logger.info(f'building 26 service values {current_accrual_values26}')
        
        #set the report month (and fill in)
        report_month_lst26 = create_report_month(month, year, current_accrual_values26)
        logger.info(f'report month: {report_month_lst26[0]}')

        #set the report address (and fill in)
        address_lst26 = create_report_address(current_accrual_values26, address_cell26)
        logger.info(f'report address: {address_lst26[0]}')

        #add report results for building 26 to the ordered dict
        for d in report_month_lst26:
            report['Дата'].append(d)
        
        for a in address_lst26:
            report['Адрес'].append(a)

        for c in accrual_names26:
            report['Атрибут'].append(c)

        for v in current_accrual_values26:
            report['Текущие Начисления'].append(v)



        ###REPORTS FOR BUILDING 30
        #get accruals for building 30
        #> an assumption that the services do not exist 25 rows
        accrual_names30 = get_accrual_names(sheet, store_row+3, (store_row+3+25), 1)
        logger.info(f'building 30 services > before removing unnecessary elements {accrual_names30}')

        #remove unnecessary elements from the list
        remove_items = []

        for e in accrual_names30:
            if e.startswith('Итог'):
                remove_items.append(e)
            elif ('в т.ч.') in e:
                remove_items.append(e)
            elif e.startswith('Пени'):
                remove_items.append(e)

        for r in remove_items:
            accrual_names30.remove(r)
            continue

        logger.info(f'building 30 services {accrual_names30}')
        
        #find the number of accruals for building 30
        length_of_names = len(accrual_names30)

        #get accrual values for building 30
        current_accrual_values30 = get_accrual_values(sheet, store_row+3, (store_row+2+length_of_names), 5, 6)
        
        logger.info(f'building 30 service values {current_accrual_values30}')
        
        #set the report month (and fill in)
        report_month_lst30 = create_report_month(month, year, current_accrual_values30)
        logger.info(f'report month: {report_month_lst30[0]}')

        #set the report address (and fill in)
        address_lst30 = create_report_address(current_accrual_values30, address_cell30)
        logger.info(f'report address: {address_lst30[0]}')

        #add report results for building 30 to the ordered dict
        for d in report_month_lst30:
            report['Дата'].append(d)
        
        for a in address_lst30:
            report['Адрес'].append(a)

        for c in accrual_names30:
            report['Атрибут'].append(c)

        for v in current_accrual_values30:
            report['Текущие Начисления'].append(v)


#write the report results to a dataframe
report_accruals_summary = pd.DataFrame(report)
logger.success(f'accrual report for {set_the_month}-{set_the_year} was created')





## MAPPING
mapping_folder_path = credentials['export_folder_path']
mapping_file_name = 'Выручка по Паркингам и кладовкам - Ведомости.xlsx'

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
reports_address = report_accruals_summary.merge(address_table,
                                                how='left',
                                                left_on='Адрес',
                                                right_on='Адрес в отчете'
                                                )

#merge the accrual report and the address map with nomenclature map
reports_address_nomenclature = reports_address.merge(nomenclature_table,
                                                     how='left',
                                                     left_on='Атрибут',
                                                     right_on='Услуга в отчете'
                                                     )

#check for duplicate values
dupl_columns = list(reports_address_nomenclature.columns)

mask = reports_address_nomenclature.duplicated(subset=dupl_columns)
duplicates = reports_address_nomenclature[mask]
logger.info(f'Number of Duplicates in Merged Tables: {duplicates.shape[0]}')

#drop duplicates (if there are any)
reports_address_nomenclature = reports_address_nomenclature.drop_duplicates(subset=dupl_columns)


#rename the columns
reports_address_nomenclature = reports_address_nomenclature.rename(columns={'Дата': 'Период'
                                                                            })


#apply the function
reports_address_nomenclature['Содержание'] = reports_address_nomenclature.apply(create_content,
                                                                                axis=1)

#add a column
reports_address_nomenclature['Кол-во'] = 1

reports_address_nomenclature['Код адреса'] = reports_address_nomenclature['Код адреса'].apply(lambda x: '000000' + str(x))


## RE-ORDER COLUMNS
#find the index of each column
idx = [reports_address_nomenclature.columns.get_loc(c) for c in reports_address_nomenclature.columns if c in reports_address_nomenclature]

#re-order columns using iloc
reports_reordered = reports_address_nomenclature.iloc[:, [0, 7, 8, 9, 10, 11, 12, 5, 6, 14, 15, 19, 16, 17, 18, 20, 3]]


#extract the columns
clms = reports_reordered.columns.to_list()

#excluding the last column
clms = clms[:-1]

#group by all > to merge the amounts by nomenclature
reports_reordered = reports_reordered.groupby(clms)['Текущие Начисления'].sum().to_frame()

#create the index > to fill down the values (repeats / duplicates)
reports_reordered = reports_reordered.reset_index()


def update_current_parking_and_storage():

    '''
        > export the current monthly results, re-write the existing results
    '''
    
    xlsx_name = 'Выручка по Паркингам и кладовкам - Ведомости.xlsx'
    sheet = 'для_загрузки'

    export_folder_path = credentials['export_folder_path']

    with pd.ExcelWriter(export_folder_path + "/" + xlsx_name,
                        engine='openpyxl',
                        mode='a',
                        if_sheet_exists='replace') as writer:
            reports_reordered.to_excel(writer,
                                       sheet_name=sheet,
                                       index=False,
                                       header=True)

    logger.success(f'new report for {set_the_month}-{set_the_year} is generated')



def create_monthly_parking_and_storage():

    '''
        > export / archive the current monthly results
    '''
    
    xlsx_name = 'Выручка по Паркингам и кладовкам - Ведомости - Архив.xlsx'
    sheet = f'для_загрузки_{set_the_month}-{set_the_year}'

    export_folder_path = credentials['export_folder_path']

    with pd.ExcelWriter(export_folder_path + "/" + xlsx_name,
                        engine='openpyxl',
                        mode='a',
                        if_sheet_exists='replace') as writer:
            reports_reordered.to_excel(writer,
                                       sheet_name=sheet,
                                       index=False,
                                       header=True)

    logger.success(f'new historical report for {set_the_month}-{set_the_year} is generated')

update_current_parking_and_storage()
create_monthly_parking_and_storage()