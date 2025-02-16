from create_the_statements import update_current_parking_and_storage
from create_the_statements import create_monthly_parking_and_storage

from create_the_untransmitted import update_current_data_for_untransmitted
from create_the_untransmitted import create_monthly_data_for_untransmitted

import os



step = os.environ['step']

def main():
    if step == 'create_table_on_statements': #создать таблицу на ведомости за текущий месяц
        update_current_parking_and_storage()
        create_monthly_parking_and_storage()

    elif step == 'create_table_on_untransmitted': #создать таблицу на непереданные данные за текущий месяц
        update_current_data_for_untransmitted()
        create_monthly_data_for_untransmitted()

if __name__ == '__main__':
    main()