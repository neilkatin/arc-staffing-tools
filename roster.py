#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime

import init_logging

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.worksheet.table as table
import openpyxl.styles.colors
import dotenv
import xlrd

import config as config_static


log = logging.getLogger(__name__)


DATESTAMP = datetime.datetime.now().strftime("%Y-%m-%d")

class AttrDict(dict):
    def __init__(self, *args, **kwargs):
        super(AttrDict, self).__init__(*args, **kwargs)
        self.__dict__ = self

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config_dotenv = dotenv.dotenv_values(verbose=True)

    config = AttrDict()
    for item in dir(config_static):
        if not item.startswith('__'):
            config[item] = getattr(config_static, item)

    #log.debug(f"config after copy: { config.keys() }")

    for key, val in config_dotenv.items():
        config[key] = val

    roster_file = config.ROSTER_FILE
    if not os.path.exists(roster_file):
        log.fatal(f"Roster file { roster_file } not found")
        sys.exit(1)

    roster_wb = xlrd.open_workbook(roster_file)
    roster_ws = roster_wb.sheet_by_index(0)

    nrows = roster_ws.nrows
    ncols = roster_ws.ncols

    #log.debug(f"roster nrows { nrows } ncols { ncols }")

    current_row = config.ROSTER_TITLE_ROW

    title_names, title_cols = parse_title_row(roster_ws, current_row)

    sup_dict, no_sups, name_dict = process_roster(roster_ws, current_row, title_names, title_cols)

    output_file = config.OUTPUT_FILE
    if os.path.exists(output_file):
        os.remove(output_file)

    output_wb = openpyxl.Workbook()


    # add the reporting info
    reporting_ws = output_wb.create_sheet(title=config.OUTPUT_SHEET_REPORTING)
    generate_sups(reporting_ws, sup_dict, name_dict, title_names)

    # add the no_sups folks
    no_sups_ws = output_wb.create_sheet(title=config.OUTPUT_SHEET_NOSUPS)
    generate_no_sups(no_sups_ws, no_sups, title_cols)


    default_sheet_name = 'Sheet'
    if default_sheet_name in output_wb:
        del output_wb[default_sheet_name]

    output_wb.save(output_file)




def parse_title_row(ws, title_row):
    """ grab all the column headers from the title row

        turn them into two dicts: column name > column number and column number > column name
    """
    title_name_dict = {}
    title_col_dict = {}

    row = ws.row_values(title_row -1)

    index = 0
    for value in row:
        #log.debug(f"title { index } = '{ value }'")
        title_name_dict[value] = index
        title_col_dict[index] = value
        index += 1

    return title_name_dict, title_col_dict


def process_roster(ws, start_row, title_names, title_cols):
    """ process all the rows in the worksheet

        ws - the worksheet
        start_row - starting row of data in the ws (origin zero)
        title_names - dict mapping names to column numbers (origin zero)
        title_cols - dict mapping column numbers to column names (origin zero)
    """

    no_sups = []
    sup_dict = {}   # keyed by sup name; contains an array of reporting people
    name_dict = {}  # keyed by name; contains whole list

    nrows = ws.nrows
    row_num = start_row -1 # convert to origin zero
    while row_num < nrows -1:
        row_num += 1

        row = ws.row_values(row_num)
        row_dict = row_to_dict(row, title_cols)

        #log.debug(f"row { row_num } values { row_dict }")

        name = row_dict['Name']
        gap = row_dict['GAP(s)']
        sup = row_dict['Current/Last Supervisor']
        released = row_dict['Released']

        name_dict[name] = row_dict

        # handle folks no longer on the job
        if released != '':
            #log.debug(f"row { row_num } person { name } has been released; skipping...")
            continue


        # handle empty supervisor field
        if sup == "":
            if gap != 'OM//DIR' and gap != 'OM//RCCO':
                no_sups.append(row)
                #log.debug(f"row { row_num } person { name } has no supervisor")
            continue

        # handle first entry in list
        if sup not in sup_dict:
            sup_dict[sup] = []

        sup_dict[sup].append(row_dict)

        # DEBUG ONLY
        #if row_num >= 180:
        #    break

    return sup_dict, no_sups, name_dict


def excel_date_to_string(excel_date):
    """ convert floating excel representation of date to yyyy-mm-dd string """

    dt = datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(excel_date) -2)

    return dt.strftime('%Y-%m-%d')



def format_reports(name, sup_row, name_dict):
    """ format the list of direct reports """

    output = [ f"{'Name':25s} { 'GAP':15s} { 'Checked in':12s} { 'Last Day':12s} {'Cell Phone':14s} { 'Email':30s}" ]

    gap_pattern = re.compile('.*,')

    for row in sup_row:

        log.debug(f"name { row['Name'] } released: { row['Released'] }")
        # only add folks still on the DR
        gap = row['GAP(s)']
        gap = re.sub(gap_pattern,'', gap)
        checkin = excel_date_to_string(row['Checked in'])
        lastday = excel_date_to_string(row['Expect release'])
        output.append(f"{ row['Name']:25s} { gap:15s} { checkin:12s} { lastday:12s} { row['Cell phone']:14s} { row['Email']:30s}")

    return "\r\n".join(output)


first_regex = re.compile('.*, ')
def get_first(s):
    return re.sub(first_regex, '', s)

last_regex = re.compile(',.*')
def get_last(s):
    return re.sub(last_regex, '', s)


def generate_sups(ws, sups, name_dict, title_dict):
    """ generate a row for each supervisor, with their direct reports and contact info """

    output_cols = [
            { 'name': 'Name', 'width': 20, 'field': lambda x: name_dict[x]['Name'], },
            { 'name': 'First', 'width': 20, 'field': lambda x: get_first(name_dict[x]['Name']), },
            { 'name': 'Last', 'width': 20, 'field': lambda x: get_last(name_dict[x]['Name']), },
            { 'name': 'Email', 'width': 30, 'field': lambda x: name_dict[x]['Email'] },
            { 'name': 'Supervisor', 'width': 30, 'field': lambda x: name_dict[x]['Current/Last Supervisor'] },
            { 'name': 'Last Day', 'width': 30, 'field': lambda x: excel_date_to_string(name_dict[x]['Expect release']) },
            { 'name': 'Reports', 'width': 50, 'field': lambda x: format_reports(x, sups[x], name_dict), },
            ]

    #log.debug(f"name_dict { name_dict.keys() }")

    # generate the title row
    for i, col_def in enumerate(output_cols):
        ws.cell(column=i+1, row=1, value=col_def['name'])
        width = col_def['width']
        #log.debug(f"column { col_def['name'] } has width { width }")
        ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = col_def['width']


    output_row = 1
    for sup_name in sorted(sups.keys()):
        output_row += 1
        sup_row = sups[sup_name]

        #log.debug(f"processing sup { sup_name } len { len(sup_row) }")

        for i, col_def in enumerate(output_cols):
            value = col_def['field'](sup_name)
            ws.cell(column=i+1, row=output_row, value=value)

    ref = f"A1:{openpyxl.utils.get_column_letter(len(output_cols))}{ output_row }"
    #log.debug(f"table range ref is '{ ref }'")
    tab = table.Table(displayName='Reports', ref=ref)
    ws.add_table(tab)


def generate_no_sups(ws, no_sups, col_dict):
    """ add the no_sups info to the worksheet

        make sure to convert from origin 0 input to origin 1 output

    """

    date_column_names = { 'Assigned':1, 'Checked in': 1, 'Expected Release': 1, }
    date_column_cols = {}

    #log.debug(f"col_dict { col_dict }")

    # convert dicts back to array
    column_array = []
    for key in sorted(col_dict.keys()):
        name = col_dict[key]
        if name in date_column_names:
            date_column_cols[key] = 1
        column_array.append(name)

    # fill in the title row
    for col, val in enumerate(column_array):
        ws.cell(column=col+1, row=1, value=val)
        #log.debug(f"setting title col { col + 1 } to { val }")

    row_num = 1
    for row in no_sups:
        row_num += 1
        col_num = -1
        for col in row:
            col_num += 1
            out_cell = ws.cell(column=col_num + 1, row=row_num, value=col)

            # make sure date columns are formatted properly
            if col_num in date_column_cols:
                out_cell.number_format = 'yyyy-mm-dd'



def row_to_dict(row, title_cols):
    """ convert a row of data into a dict mapping column name to value """

    result = {}
    for index, value in enumerate(row):
        result[title_cols[index]] = value

    return result

def gather_column(ws, column_num, title_row_num, table_name, column_name):
    """ gather all the values in a column into dict for easy matching """

    results = {}
    row_num = title_row_num
    for row in ws.iter_rows(min_row=title_row_num+1, min_col=column_num, max_col=column_num, values_only=True):
        row_num += 1
        cell = row[0]
        if cell == '':
            continue

        # delete formatting characters
        if isinstance(cell, str):
            cell = re.sub('[ \t-]', '', cell)

        value = str(cell)
        if value in results:

            # horrible hack to deal with n/a column in reservation no: just ignore values of n/a
            if value != 'N/A':
                # this is a duplicate entry
                log.error(f"Found duplicate for table { table_name } column { column_name }: old row { results[value] }, new row { row_num }")

        results[value] = row_num

    return results

def process_title_row(sheet, title_row_num):
    """ build two dicts: one mapping column name to column number, and one from column number to name """
    title_name_map = {}
    title_cols = {}

    for row in sheet.iter_rows(min_row=title_row_num, max_row=title_row_num, values_only=False):
        for cell in row:
            #log.debug(f"title { cell.column }, { cell.row } = '{ cell.value }'")
            title_name_map[cell.value] = cell.column
            title_cols[cell.column] = cell.value
        break

    return title_name_map, title_cols

def build_map(sheet, sheet_name, starting_row, key_name_list, row_filter):
    """ build a dict of all values in a sheet, keyed by the given column

        use key_name_list to find the key for each entry in the returned dict.  The first
        non-empty value is the list of column names will be the key.

        The value for each dict entry will in turn be a dict of all the column-name to value cells in that row.
        The dict entry will have two additional values added:

          row_num: the number of the row (origin 1 in the sheet) the value appeared in
          sheet_name: the sheet name (or label) passed into this function
    """

    title_name_map, title_cols = process_title_row(sheet, starting_row)

    #log.debug(f"title_name_map: { title_name_map }")

    results = {}
    row_num = starting_row
    for row in sheet.iter_rows(min_row=starting_row+1, values_only=False):
        row_num += 1
        row_map = { 'row_num': row_num, 'sheet_name': sheet_name }
        for cell in row:
            value = cell.value
            title = title_cols[cell.column]
            row_map[title] = value

        if not row_filter(row_map):
            continue

        #log.debug(f"row_map: { row_map }")
        for name in key_name_list:
            key = row_map[name]
            if key != '' and key != None:
                break

        if key == '' or key == None:
            log.error(f"build_map: empty key from { key_name_list } for row { row_num }")

        if key in results:
            log.error(f"Error: duplicate entry for entry { key } on row { row_num } and { results[key]['row_num'] }")

        results[key] = row_map

    return results


def parse_args():
    parser = argparse.ArgumentParser(
            description="process support for the regional bootcamp mission card system",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")

    args = parser.parse_args()
    return args


if __name__ == "__main__":
    main()

