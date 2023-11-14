import pandas as pd
from datetime import datetime, timedelta, time
from dateutil import rrule
import sys
import PyPDF2 as pypdf
import calendar
import re
from tqdm import tqdm
from pprint import pprint
from config import get_config


"""=== Functions for statistical analysis ==="""


def time_diff(record) -> float:
    """ The function is considering the time interval: Start-End
    Start and End has the format: HH:MM
    :return hours as float rounded to 2 digits after period
    Example:
        8:30-12:15
        Difference = 3.75
    """
    start: time = record['Start']
    end: time = record['End']
    date: datetime = record['Date']
    # convert datetime.time to datetime.datetime
    start_date = datetime(hour=start.hour, minute=start.minute, year=date.year, month=date.month, day=date.day)
    end_date = datetime(hour=end.hour, minute=end.minute, year=date.year, month=date.month, day=date.day)
    difference = (end_date - start_date).total_seconds()/(60*60)
    difference = round(difference, 2)
    return difference


def get_students_num(record) -> int:
    """Function tries to find the column which should include the number of the students in the venue"""
    column_name: str = get_config()['input_student_num_col']
    if len(column_name) == 0:
        return 0
    column_name = column_name.strip('()')
    num = int(record[column_name])
    return num


"""=== Functions for formatting dataframes ==="""


def merge_lists_to_second(list1: list, list2: list) -> None:
    for i in list1:
        if i not in list2:
            list2.append(i)


def check_format(df: pd.DataFrame, required_cols: list[str]) -> None:
    """
    The function is determining whether the original dataframe contain necessary columns
    Headers of the required columns are stored in the array and parsed one by one.
    """
    df_cols: list = df.columns
    for col in required_cols:
        if col not in df_cols:
            print(f'Column "{col}" is missing in the DataFrame', file=sys.stderr)
            break


def split_column(df: pd.DataFrame, column_name: str, new_column_name: str, delimiter='') -> pd.DataFrame:
    """
    The function is copying the column and appends to the main dataframe
    """
    columns_to_extract = [col for col in df.columns if col != column_name]
    # create new column
    new_column_series: pd.Series = df.apply(lambda row: row[column_name].split(delimiter)[1], axis=1)
    new_column_series.rename(new_column_name, inplace=True)
    # update old column
    date_format = "%Y-%m-%d %H:%M:%S"
    column_series: pd.Series = df.apply(lambda row: row[column_name].split(delimiter)[0], axis=1)
    column_series = column_series.apply(lambda value: datetime.strptime(value, date_format))
    column_series.rename(column_name, inplace=True)
    # add new columns to the passed dataframe
    df = df[columns_to_extract]
    df = df.join(new_column_series)
    df = df.join(column_series)
    return df


def formatted_booking(record, time_format: str, target_column: str, columns: list[str]):
    config_key = 'booking_format_'+target_column.lower()
    form = get_config()[config_key]
    for key in columns:
        value = record[key]
        if type(value) != str:
            value = value.strftime(time_format)
        form = form.replace(f"[{key}]", value)
    return form


def get_booking_details(df: pd.DataFrame, target_column: str, delimiter='|') -> pd.DataFrame:
    """
    The function is setting the new data cell in the format given below
    Format:

    Staff Name      | Date
    Booking detail1 | yyyy-mm-dd
    NaN             | yyyy-mm-dd
    Booking detail2 | yyyy-mm-dd

    Format of Booking detail:
    'Subject Code|hh:mm-hh:mm|taskName'

    """
    time_format = '%H:%M' # the time period is represented as hh:mm-hh:mm
    formatted_bookings = {}
    """var formatted_bookings has the following format:
    {
        'date|session': {
            staff_name1: booking_detail1,
            staff_name2: booking_detail2,
            venue1: booking_detail1
        }
    }
    NOTE: staff names and venue names are target_col variable
    """
    record_stats = {}
    """ var record_stats has the following format: 
    {
        'date|session':
        {
            staff_name1 (Hours): float
            venue1 (Hours): float
            venue1 (Students number): int
            staff_name2 (Hours): int
        }
    }
    NOTE: staff names and venue names are target_col variable
    """
    total: int = df.shape[0]
    bar_progress: tqdm = tqdm(total=total, disable=False)
    for i in range(total):
        record = df.iloc[i]
        """Unpack each record"""
        column_names: list[str] = get_config()['required_columns']
        for col in column_names:
            df[col].fillna(value='', inplace=True) # change NaN to '' string
        target_col: str         = record[target_column]
        date: datetime          = record['Date']
        session: str            = record['Start'].strftime('%p') # AM/PM session
        """ Format details """
        booking_details = formatted_booking(record,
                                            columns=df.columns,
                                            time_format=time_format,
                                            target_column=target_column)
        time_difference: float = time_diff(record) # append it to the final dataframe
        number_students: int   = get_students_num(record)
        # booking_details += delimiter + str(time_difference) + delimiter + str(number_students)
        key = str(date)+'|'+session
        """ Prepare booking details """
        entry = {target_col: booking_details}
        if key not in formatted_bookings.keys():
            formatted_bookings[key] = entry
        else:
            formatted_bookings[key].update(entry)

        """ Prepare statistics """
        entry_h = {target_col+get_config()['output_hours_num_col']: time_difference}
        entry_num = {target_col + get_config()['output_student_num_col']: number_students}
        if key not in record_stats.keys():
            record_stats[key] = entry_h
            if target_column == 'Venue':
                record_stats[key].update(entry_num)
        else:
            record_stats[key].update(entry_h)
            if target_column == 'Venue':
                record_stats[key].update(entry_num)

        bar_progress.update() # by default, updates by 1
    bar_progress.close()

    print(record_stats)

    # pprint(formatted_bookings)
    """ ==== Convert records into dataframe ==== """
    ds = pd.DataFrame.from_records(formatted_bookings)
    ds = ds.transpose()
    stats_ds = pd.DataFrame.from_records(record_stats)
    stats_ds = stats_ds.transpose()
    print('TARGET DATAFRAME')
    print(ds)
    print('STATS')
    print(stats_ds)
    ds: pd.DataFrame = pd.merge(ds,
                                stats_ds,
                                left_index=True,
                                right_index=True,
                                how='outer')
    print('RESULT')
    ds = ds[sort_df_columns(ds)]
    print(ds.head())
    ds.to_csv(path_or_buf=get_config()['testing_file'])
    ds.reset_index(inplace=True)
    ds.rename(columns={'index': 'Date'}, inplace=True)
    ds = split_column(df=ds, column_name='Date', new_column_name='Session', delimiter='|')
    return ds


def get_dates_in_ac(start_year: int, start_month: int, end_year=None, end_month=None, end_day=None) -> pd.DataFrame:
    """
    The function return the list of all days within an academic year
    The list represents a pd.Dataframe with the column names:
    ['Date', 'Session', 'Week', 'Day']
    """
    start_date = datetime(start_year, start_month, 1) # start with the first day of the month
    if end_year is None and end_month is None and end_day is None:
        end_date = start_date + timedelta(days=365) # end after one calendar year
    else:
        end_date = datetime(year=end_year, month=end_month, day=1)
    """ == Define patterns == """
    date_format = "%d-%b-%y" # dd-MonthName-yy
    day_name = "%a" # Friday, Monday, etc.
    # org_date_format = '%Y-%m-%d' # time format yyyy-mm-dd used in the
    """ ==== Compress the data into list[dict] ==== """
    date_values = list(dict())
    for i in range((end_date - start_date).days):
        date = start_date + timedelta(days=i)
        """ ==== Get the formatted data ==== """
        formatted_date: str = date.strftime(date_format)
        week_number: int    = get_week_num(this_day=date, start_date=start_date) # get the number of the week relative to the start day
        week_name: str      = date.strftime(day_name) # get the name of the day
        """ ==== Compress the data into dictionary with AM and PM annotation ===== """
        sessions = ['AM', 'PM'] # there are two class sessions per day
        for each_session in sessions:
            date_values.append({'Date': date,
                                'Formatted Date': formatted_date,
                                'Session': each_session,
                                'Week': week_number,
                                'Day': week_name})
    return pd.DataFrame(data=date_values)


def get_unique_values(df: pd.DataFrame, column_name: str) -> list[str]:
    """
    Function is looking for the names in the dataframe and
    extract the unique names with sorting
    """
    df = df[~df[column_name].isna()] # ignore the NaN
    # Alternative method:
    # unique_values = list(df.column_name.unique())
    # unique_values.sort()
    # return unique_values
    staff_list = list(df[column_name])
    staff_list = list(set(staff_list))  # unique get unique values from the list
    staff_list.sort()
    return staff_list


def get_week_num(this_day: datetime, start_date: datetime) -> int:
    """
    The function return the week number relative to the starting date
    If start date == 01 Sept and this day == 02 Sept => return 1
    It is determined by the number of weeks between this day and the starting day
    """
    weeks = rrule.rrule(rrule.WEEKLY, dtstart=start_date, until=this_day)
    return weeks.count()


def extract_dates(text: str, included_classes_suspended=False) -> dict:
    """
    The text from the pdf is split into lines and keyword "General holiday" and "suspended" is matching.
    The function return the dictionary, containing the number of days per holiday and name of holiday
    :returns {'holiday name': [1 September 2023, 2 September 2033]}
    """
    month_names: list[str] = [calendar.month_name[num] for num in range(1, 13)]
    pattern = r"\b\d{4}\b"
    match = re.search(pattern, text)
    year = str(match.group())
    text: list[str] = text.split('\n') # split the text into lines
    this_month = '' # store the name of the month
    result = {}
    for line in text:
        line = line.strip()
        if len(line.split(' ')) == 1 and line in month_names:
            # it means the line consists of only one word, which indicates the name of the month.
            this_month = line
        if line.__contains__('General holiday'):
            # extracting name of the holiday from the line
            holiday_name: str = re.findall(pattern=r"\((.*?)\)", string=line)[0] # excepted that each line = one holiday
            holiday_dates: list[str] = re.findall(pattern=r"\d+", string=line)
            holiday_dates: list[str] = [str(holiday_date)+' '+this_month+' '+year for holiday_date in holiday_dates]
            result[holiday_name] = holiday_dates
        elif 'classes' in line and ':' in line and included_classes_suspended:
            name_date: str = line.split('(')[0]
            name_date = name_date.replace(' - ', '-')
            holiday_name = name_date.split(':')[1].strip()
            properties = name_date.replace(':', '').split(' ')
            days, month = properties[0], properties[1]
            month = month[:3] # extract only the first 3 letters of the month name
            if not days.__contains__('-'):
                this_date = datetime(year=int(year), month=datetime.strptime(month, "%b").month, day=int(days))
                formatted_date = f"{this_date.day} {this_date.strftime('%B')} {year}"
                result[holiday_name] = [formatted_date]
            else:
                start_date, end_date = map(int, days.split("-"))
                this_date = datetime(year=int(year), month=datetime.strptime(month, "%b").month, day=start_date)
                day_index = 0
                while this_date.day <= end_date:
                    day_index += 1
                    formatted_date = f"{this_date.day} {this_date.strftime('%B')} {year}"
                    result[holiday_name+str(day_index)] = [formatted_date]
                    this_date += timedelta(days=1)
    return result


def get_holidays(ac_path: str, pdf_pages) -> pd.DataFrame:
    """
    Function returns the dataframe with public holidays in Hong Kong
    The format:
    Date        | Description |
    yyyy-mm-dd  |   str       |
    """
    """ ===== Get the page content as text ===== """
    pdf_file = open(ac_path, 'rb')  # r = read string, rb = read binary
    pdf_file_obj: pypdf.PdfReader = pypdf.PdfReader(pdf_file)  # PdfFileReader is not available
    pdf_page_content = ''
    if pdf_pages == 'all':
        pdf_pages = [i for i in range(len(pdf_file_obj.pages))]
    assert type(pdf_pages) == list
    for i in pdf_pages:
        pdf_page_content += pdf_file_obj.pages[i].extract_text()
    # print(pdf_page_content)
    general_holidays: dict = extract_dates(text=pdf_page_content, included_classes_suspended=False)
    """ ==== Convert and format the extracted days into a list ==== """
    df: pd.DataFrame = pd.DataFrame.from_dict(data=general_holidays) # general holidays dataframe
    df = df.transpose()
    df.reset_index(inplace=True)
    df.rename(columns={'index': 'Description', 0: 'Date'}, inplace=True)
    format_string = "%d %B %Y"
    df['Date'] = pd.to_datetime(df["Date"], format=format_string)
    return df


def sort_df_columns(df: pd.DataFrame) -> list:
    """ Sort the column names according to the format:
    ['Venue1', 'Venue1(Hours)', 'Venue1(Student Number)', 'Venue2', 'Venue2(Hours)', ...]
    """
    # get the unique names of columns excluding content of ()
    cols: list[str] = [col_name for col_name in df.columns if '(' not in col_name]
    cols_sorted = []
    for col_name in cols:
        cols_sorted.append(col_name)
        col_name_h = col_name+get_config()['output_hours_num_col']
        if col_name_h in df.columns:
            cols_sorted.append(col_name_h)
        col_name_num = col_name+get_config()['output_student_num_col']
        if col_name_num in df.columns:
            cols_sorted.append(col_name_num)
    return cols_sorted


def get_up_to_date(df: pd.DataFrame, day=1, month=get_config()['start_month'], year=get_config()['start_year']) -> pd.DataFrame:
    date: str = datetime(day=day, month=month, year=year).strftime("%d-%b-%y")
    return df[:date]
