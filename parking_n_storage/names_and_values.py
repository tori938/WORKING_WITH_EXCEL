def convert_address(element: str):

    '''
        > convert the address line to address line in 1C
    '''

    if element == 'г. Санкт-Петербург ш Суздальское, д. 30 к.2 стр.1':
        return 'Суздальское д. 30 к.2 стр.1'
    elif element == 'г. Санкт-Петербург ш Суздальское, д. 26 к.1':
        return 'Суздальское д. 26 к.1'



def convert_month_to_number(text: str):

    '''
        > convert month to a numerical value (1-12)
    '''
    if text == 'Январь':
        return 1
    elif text == 'Февраль':
        return 2
    elif text == 'Март':
        return 3
    elif text == 'Апрель':
        return 4
    elif text == 'Май':
        return 5
    elif text == 'Июнь':
        return 6
    elif text == 'Июль':
        return 7
    elif text == 'Август':
        return 8
    elif text == 'Сентябрь':
        return 9
    elif text == 'Октябрь':
        return 10
    elif text == 'Ноябрь':
        return 11
    elif text == 'Декабрь':
        return 12
    else:
        return 0



def get_accrual_names(active_sheet,
                      row_start: int,
                      row_end: int,
                      column_index: int):
    '''
        > extract the names of accruals
    '''

    accrual_names = list()
    
    for i in range(row_start-1, row_end):
        for j in range(column_index):
            #take the elements from specific rows (i.e. 5-19 for b26) and 1st column
            accrual_cell_value = active_sheet.cell_value(i, j)

            #remove spaces from the beginning and end of the element only
            accrual_names.append(accrual_cell_value.lstrip().rstrip())
    
    return accrual_names



def get_accrual_values(active_sheet,
                       row_start: int,
                       row_end: int,
                       column_idx_start: int,
                       column_idx_end: int):
    '''
        > extract the accrual values
    '''

    accrual_values = list()
    
    for i in range(row_start-1, row_end):
        for j in range(column_idx_start, column_idx_end):

            #take the elements from specific rows (i.e. 5-19 for b26) and 5th column
            accrual_value = active_sheet.cell_value(i, j)
            
            #to account for export errors: string > float
            if type(accrual_value) == str:
                accrual_values.append(float(accrual_value.replace(",", ".")))
            else:
                accrual_values.append(accrual_value)
    
    return accrual_values



def get_cost_names(active_sheet,
                   row_idx: int,
                   column_idx: int,
                   total_columns: int):
    
    '''
        > extract the names of deductions 
    '''

    cost_names = list()
    
    for i in range(row_idx):
        for j in range(column_idx, total_columns):

            #if there are spaces between columns, do not take the column value
            if active_sheet.cell_value(i,j) != '':
                cost_cell = active_sheet.cell_value(i,j)

                #remove spaces from the beginning and end of the element only
                cost_names.append(cost_cell.lstrip().rstrip())
    
    return cost_names



def get_cost_values(active_sheet,
                    total_rows: int,
                    column_idx: int,
                    total_columns: int):
    
    '''
        > extract the deduction values
    '''

    cost_values = list()
    
    for i in range((total_rows-1), total_rows):
        for j in range(column_idx, total_columns):
            #if there are spaces between columns, do not take the column value
            if active_sheet.cell_value(i,j) != '':
                cost_value = active_sheet.cell_value(i,j)

                #if the value is not a number
                if type(cost_value) == str:
                    cost_value = float(cost_value.replace(",", "."))
                    #if the value is less than 0 (or negative) > new update: do not change
                    if cost_value < 0:
                        #append the list with a zero value
                        #cost_values.append(0)
                        cost_values.append(cost_value)
                    else:
                        cost_values.append(cost_value)
                else:
                    #if the value is less than 0 (or negative) > new update: do not change
                    if cost_value < 0:
                        #append the list with a zero value
                        #cost_values.append(0)
                        cost_values.append(cost_value)
                    else:
                        cost_values.append(cost_value)
    
    return cost_values



def create_report_month(m: int,
                        y: int,
                        the_list_of_accrual_values: list):
    
    '''
        > create duplicate values for the month to fill up in the report
    '''

    #add 0 to single-digit months
    if len(str(m)) == 1:
        m = '0' + str(m)

    #create a report date
    report_date = '01.' + str(m) + '.' + str(y)

    #create an empty list to store the report month (duplicated)
    report_lst = []

    length = len(the_list_of_accrual_values)
    i = 0

    #fill in the report date based on the length of the list of values
    while i < length:
        report_lst.append(report_date)
        i += 1
    
    return report_lst



def create_report_address(the_list_of_accrual_values: list,
                          address_name: str):

    '''
        > create duplicate values for the address line to fill up in the report
    '''

    #create an empty list to store the address (duplicated)
    address_lst = []

    length = len(the_list_of_accrual_values)
    i = 0

    #fill in the address based on the length of the list of values
    while i < length:
        address_lst.append(address_name)
        i += 1
    
    return address_lst



def create_content(data):

    '''
        > convert the content to match nomenclature change
    '''
    
    if data['Номенклатура 1С'] == 'ГВС компонент ТН в целях СОИ = ОДН ГВС':
        return 'ОДН ГВС'
    elif data['Номенклатура 1С'] == 'ГВС компонент ТЭ в целях СОИ = ОДН ГВС':
        return 'ОДН ГВС'
    elif data['Номенклатура 1С'] == 'Водотвед.гор.(общед.нужды) = ОДН ВО':
        return 'ОДН ВО'
    else:
        return data['Номенклатура 1С']