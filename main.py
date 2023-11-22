import openpyxl # !important for reading excel files!
from openpyxl import Workbook
import data_processing as dp
import excel_style as es
import pandas as pd
from config import get_config

""" ===== Define configuration ===== """
config = get_config()
start_month: int                    = config['start_month']
start_year: int                     = config['start_year']
weekends: list[str]                 = config['weekends']
input_file_path: str                = config['input']
output_file_path: str               = config['output']
ac_file_path: str                   = config['ac_file_path'] # academic calendar path file
target_columns: list[str]           = config['target_columns']
required_cols: list                 = config['required_columns']
pdf_pages                           = config['ac_pdf_pages'] # can be list[int] or 'all'
fixed_columns: int                  = config['freeze_columns']
fixed_rows: int                     = config['freeze_rows']
output_hours_num_col: str           = config['output_hours_num_col']
output_student_num_col: str         = config['output_student_num_col']
dp.merge_lists_to_second(list1=target_columns, list2=required_cols)

if __name__ == '__main__':
    """ ===== Import original data ===== """
    df_org: pd.DataFrame = pd.read_excel(input_file_path, header=0)  # get original data
    df_org = df_org.iloc[:len(df_org) - 1, :]  # remove the last unnecessary row
    print(df_org)
    """ === Get Academic calendar === """
    ac_df: pd.DataFrame = dp.get_dates_in_ac(start_year=start_year,
                                             start_month=start_month,
                                             end_year=start_year+1,
                                             end_month=12,
                                             end_day=1)  # get academic calendar
    holidays: pd.DataFrame = dp.get_holidays(ac_path=ac_file_path, pdf_pages=pdf_pages)  # get general holidays calendar
    print(holidays)
    """ === Copy the original dataframe === """
    dp.check_format(df_org, required_cols)
    """ === Analyse data === """
    with pd.ExcelWriter(output_file_path) as writer:
        for target_column in target_columns:
            """ ===== Format the booking details ===== """
            df: pd.DataFrame = df_org[required_cols]  # get the dataframe with required columns only
            df = df.dropna(subset=[target_column])  # remove the records with staff = nan
            print('=' * 5 + f'PROCESSING DATA: "{target_column.upper()}" WORKSHEET' + '=' * 5)
            df[required_cols].fillna(value='', inplace=True)  # remove the nan values from column Task
            df_bookings: pd.DataFrame = dp.get_booking_details(df=df, target_column=target_column, delimiter='\n')
            # print(df_bookings)
            """ ===== Combine academic calendar and booking dataframes ===== """
            df: pd.DataFrame = pd.merge(ac_df,
                                        df_bookings,
                                        left_on=['Date', 'Session'],
                                        right_on=['Date', 'Session'],
                                        how='outer')
            df.drop_duplicates(inplace=True)

            unique_names = dp.get_unique_values(df_org, column_name=target_column)
            cols = dp.sort_df_columns(df)
            # exclude column names which does not have target column values
            cols = [col for col in cols if col in unique_names
                    or col.__contains__(output_hours_num_col.strip('()'))
                    or col.__contains__(output_student_num_col.strip('()'))]
            print('Columns found')
            print(cols)

            """=== Add total into column name ==="""
            new_column_names = {}
            for col in df.columns:
                if output_hours_num_col not in col and output_student_num_col not in col:
                    continue
                total: str = str(int(df[col].sum(skipna=True)))
                # reformat the column name to "Venue1(Hours:total)" and "Venue1(Student Number:total)"
                previous_col = col
                col = col[:len(col) - 1] + ':' + total + col[len(col) - 1]
                # store previous column name and new column name
                new_column_names[previous_col] = col
            print('NEW COLUMN NAMES')
            print(new_column_names)

            """ ===== Indicate the holidays on the main dataframe ===== """
            df.loc[df['Day'].isin(weekends), cols] = 'PH'
            df.loc[df['Date'].isin(dp.get_unique_values(holidays, column_name='Date')), cols] = 'PH'

            """ ===== Save total number to column names ===="""
            df.rename(columns=new_column_names, inplace=True)

            """ ==== Save the dataframe and set index ==== """
            extract_columns = [col for col in df.columns if col != 'Date']
            df = df[extract_columns]
            df.rename(columns={'Formatted Date': 'Date'}, inplace=True)
            df.dropna(subset=['Date'], inplace=True)
            df.set_index('Date', inplace=True)
            print(df)
            print(f'{df.shape[0] - ac_df.shape[0]} new rows were added')

            """ ==== Show up to certain day ==== """
            df = dp.get_up_to_date(df, month=2, year=2024)

            """ ==== Add the dataframe to workbook ==== """
            # df.to_csv('data_storage/test.csv')  # save the dataframe in csv file
            df.to_excel(writer, sheet_name=target_column)

    """ ==== Change excel style ==== """
    for worksheet_name in target_columns:
        workbook: Workbook = openpyxl.load_workbook(filename=output_file_path, read_only=False)
        ws = workbook[worksheet_name]
        es.adjust_text_alignment(worksheet=ws)
        # autoresize all columns
        es.autoresize_columns(worksheet=ws)
        es.freeze(worksheet=ws, columns=fixed_columns, rows=fixed_rows)
        # autoresize freeze columns
        # es.autoresize_columns(worksheet=ws, starting_column=1, ending_column=fixed_columns, column_width=10)
        workbook.save(filename=output_file_path)
        print(10 * '=' + 'NEW CLASS TIMETABLE HAS BEEN SAVED' + '=' * 10)
