#! /usr/bin/env python3

# gaptastic -- match open door positions to available responders

import argparse
import logging
import os
import os.path
import re
import datetime
import time
import pprint
import json
import io
import csv
import sys
import random

import requests
import requests_html
import dotenv
import xlrd
import openpyxl

from http.cookiejar import LWPCookieJar, Cookie


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import init_logging
from vc_session import get_session
log = logging.getLogger(__name__)
import config as config_static


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = init_config()

    session = get_session(config)
    #process_arrival_roster(read_arrival_roster(session, config, False))

    #with open('arrival_roster.xls', 'rb') as f:
    #    contents = f.read()
    #    process_arrival_roster(contents)

    #with open('Open Staff Requests.xls', 'rb') as f:
    #    contents = f.read()
    #    process_open_requests(contents)

    with open('staff_requests.xls', 'rb') as f:
        contents = f.read()
        process_staff_roster(contents)


ORDINAL_1900_01_01 = datetime.datetime(1900, 1, 1).toordinal()
TODAY = datetime.date.today()
LEFT_ALIGN = openpyxl.styles.Alignment(horizontal='left')

def process_arrival_roster(contents):

    fill_today = openpyxl.styles.PatternFill(fgColor='C9E2B8', fill_type='solid')
    fill_tomorrow = openpyxl.styles.PatternFill(fgColor='9BC2E6', fill_type='solid')
    fill_past = openpyxl.styles.PatternFill(fgColor='FFDB69', fill_type='solid')

    def pre_fixup(in_ws, out_ws):
        # copy the title values
        out_ws['A1'] = in_ws.cell_value(0,0)
        out_ws['A2'] = in_ws.cell_value(1,0)
        out_ws['A3'] = in_ws.cell_value(2,0)

        title_font = openpyxl.styles.Font(name='Arial', size=14, bold=True)
        out_ws['k1'].font = out_ws['k2'].font = out_ws['k3'].font = title_font
        out_ws['a1'].font = out_ws['a2'].font = out_ws['a3'].font = title_font

        out_ws.cell(row=1, column=11, value='Arriving Today').fill = fill_today
        out_ws.cell(row=2, column=11, value='Arriving Tomorrow').fill = fill_tomorrow
        out_ws.cell(row=3, column=11, value='Past Due Date').fill = fill_past

    def filter_arrive_date(cell, today, fill_past, fill_today, fill_tomorrow):
        """ decide if there is a special fill to apply to the cell """
        excel_date = cell.value
        dt = datetime.datetime.fromordinal(ORDINAL_1900_01_01 + int(excel_date) -2)
        date = dt.date()

        if date < today:
            cell.fill = fill_past
        elif date == today:
            cell.fill = fill_today
        elif date == datetime.timedelta(1) + today:
            cell.fill = fill_tomorrow


    params = {
            'sheet_name': 'Arrival Roster',
            'out_file_name': 'arrival_roster.xlsx',
            'table_name': 'Arrival',
            'in_starting_row': 5,
            'out_starting_row': 4,
            'column_formats': {
                    'Arrive date': 'yyyy-mm-dd',
                    'Flight Arrival Date/Time': 'yyyy-mm-dd HH:MM',
                    },
            'column_widths': {
                    'Name': 25,
                    'Flight Arrival Date/Time': 25,
                    'Flight City': 18,
                    'District': 18,
                    'GAP': 12,
                    'Arrive date': 14,
                    'Reporting/Work Location': 22,
                    'Email': 25,
                    'Cell phone': 13,
                    'Home phone': 13,
                    'Work phone': 13,
                    },
            'column_alignments': {
                    'Arrive date': LEFT_ALIGN,
                    'Flight Arrival Date/Time': LEFT_ALIGN,
                    },
            'column_fills': {
                    'Arrive date': lambda cell: filter_arrive_date(cell, TODAY, fill_past, fill_today, fill_tomorrow),
                    },
            'pre_fixup': pre_fixup,
            }

    return process_common(contents, params)


def process_open_requests(contents):

    def pre_fixup(in_ws, out_ws):
        # copy the title values
        out_ws['A1'] = in_ws.cell_value(0,0)
        out_ws['E1'] = in_ws.cell_value(0,3)

        date_cell = out_ws['A2']
        date_cell.value = datetime.datetime.now()
        date_cell.number_format = 'yyyy-mm-dd hh:mm'
        date_cell.alignment = LEFT_ALIGN

        title_font = openpyxl.styles.Font(name='Arial', size=14, bold=True)
        out_ws['A1'].font = out_ws['E1'].font = title_font

    def post_fixup(in_ws, out_ws, title_dict):
        out_ws.delete_cols(4, 1)

        # now fix up the column widths, since delete_cols doesn't update column_dimensions
        col_dims = out_ws.column_dimensions
        max_col = out_ws.max_column
        for col in range(5, max_col +2):
            from_col_letter = openpyxl.utils.cell.get_column_letter(col)
            to_col_letter = openpyxl.utils.cell.get_column_letter(col -1)
            col_dims[to_col_letter].width = col_dims[from_col_letter].width
            #log.debug(f"copying { from_col_letter } { col } to { to_col_letter }, width { col_dims[to_col_letter].width }")



    params = {
            'sheet_name': 'Open Staff Requests',
            'out_file_name': 'Open Staff Requests.xlsx',
            'table_name': 'OpenRequests',
            'in_starting_row': 2,
            'out_starting_row': 3,
            'column_formats': {
            #        'Arrive date': 'yyyy-mm-dd',
            #        'Flight Arrival Date/Time': 'yyyy-mm-dd HH:MM',
                    },
            'column_widths': {
                    'Proximity': 16,
                    'G/A/P': 12,
                    'Qual 1': 15,
                    'Location': 18,
                    'Supervisor': 24,
                    },
            'column_alignments': {
            #        'Arrive date': LEFT_ALIGN,
            #        'Flight Arrival Date/Time': LEFT_ALIGN,
                    },
            
            'column_fills': {
            #        'Arrive date': lambda cell: filter_arrive_date(cell, today, fill_past, fill_today, fill_tomorrow),
                    },
            'pre_fixup': pre_fixup,
            'post_fixup': post_fixup,
            }

    return process_common(contents, params)



def process_staff_roster(contents):
    """ generate the staff roster spreadsheets """

    fill_remain = openpyxl.styles.PatternFill(fgColor='FFDB69', fill_type='solid')

    def pre_fixup(in_ws, out_ws):
        # copy the title values
        out_ws['A1'] = in_ws.cell_value(0,0)
        out_ws['A2'] = in_ws.cell_value(1,0)
        out_ws['A3'] = in_ws.cell_value(2,0)

        title_font = openpyxl.styles.Font(name='Arial', size=14, bold=True)
        out_ws['A1'].font = out_ws['A2'].font = out_ws['A3'].font = title_font

    def row_filter(row, title_dict):
        released = row[title_dict['Released']]
        #log.debug(f"checking released: { released }")

        return released == ''

    def filter_days_remain(cell, today, fill_remain):

        value = cell.value
        if value != "" and value != 'n/a':
            value = cell.value = int(cell.value)
            if value <= 2:
                cell.fill = fill_remain

    def filter_on_job(cell):
        value = cell.value
        if value != "" and value != 'n/a':
            cell.value = int(cell.value)



    params = {
            'sheet_name': 'Staff Roster',
            'out_file_name': 'Staff Roster.xlsx',
            'table_name': 'StaffRoster',
            'in_starting_row': 5,
            'out_starting_row': 4,
            'column_formats': {
                    'Assigned': 'yyyy-mm-dd',
                    'Checked in': 'yyyy-mm-dd',
                    'Released': 'yyyy-mm-dd',
                    'Expect release': 'yyyy-mm-dd',
                    },
            'column_widths': {
                    'Name': 25,
                    'Assigned': 11,
                    'Checked in': 11,
                    'Released': 11,
                    'Current/Last Supervisor': 22,
                    'GAP(s)': 14,
                    'District': 16,
                    'Reporting/Work Location': 36,
                    'Location type': 16,
                    'Expect release': 11,
                    'Last action': 12,
                    'Current lodging': 20,
                    'Qualifications': 32,
                    'All GAPs': 32,
                    'Languages': 28,
                    'All Supervisors': 28,
                    'Evaluation status(es)': 28,
                    'COVID-19 issuer/ notes': 48,
                    'Email': 33,
                    'Cell phone': 14,
                    'Home phone': 14,
                    'Work phone': 14,
                    ' ZIP': 12,         # yes, there really is a space in the column title as generated by VC
                    },
            'column_alignments': {
                    'Assigned': LEFT_ALIGN,
                    'Checked in': LEFT_ALIGN,
                    'Released': LEFT_ALIGN,
                    'Expect release': LEFT_ALIGN,
                    },
            
            'column_fills': {
                    'DaysRemain': lambda cell: filter_days_remain(cell, TODAY, fill_remain),
                    'On Job': lambda cell: filter_on_job(cell)
                    },
            'pre_fixup': pre_fixup,
            #'post_fixup': post_fixup,
            'row_filter': row_filter,
            }

    process_common(contents, params)

    params['sheet_name'] = 'Outprocessed'
    params['out_file_name'] = 'Outprocessed Roster.xlsx'
    params['table_name'] = 'Outprocessed'
    params['row_filter'] = lambda row, title_dict: not row_filter(row, title_dict)

    process_common(contents, params)

def process_common(contents, params):
    """ common code to process all sheets """
    
    in_starting_row = params['in_starting_row']
    out_starting_row = params['out_starting_row']

    in_wb = xlrd.open_workbook(file_contents=contents)
    in_ws = in_wb.sheet_by_index(0)

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.create_sheet(title=params['sheet_name'])

    # do some ws dependent preliminary initialization
    if 'pre_fixup' in params:
        params['pre_fixup'](in_ws, out_ws)

    # deal with the title row
    title_dict = process_title_row(in_ws.row_values(in_starting_row))
    col_format = {}
    for col_name, cell_format in params['column_formats'].items():
        col_index = title_dict[col_name]
        col_format[col_index] = cell_format

    for col_name, col_width in params['column_widths'].items():
        col_index = title_dict[col_name]
        col_letter = openpyxl.utils.cell.get_column_letter(col_index + 1)
        out_ws.column_dimensions[col_letter].width = col_width

    col_alignment = {}
    for col_name, alignment in params['column_alignments'].items():
        col_index = title_dict[col_name]
        col_alignment[col_index] = alignment

    col_fills = {}
    for col_name, fill in params['column_fills'].items():
        col_index = title_dict[col_name]
        col_fills[col_index] = fill

    # now deal with the body of the message
    num_rows = in_ws.nrows - in_starting_row
    row_filter = lambda x, y: True
    if 'row_filter' in params:
        row_filter = params['row_filter']

    log.debug(f"num_rows { num_rows }")
    out_index = out_starting_row -1
    for index in range(num_rows):
        in_row = in_ws.row_values(index + in_starting_row)

        # allow us to ignore rows
        if index > 0 and not row_filter(in_row, title_dict):
            continue

        out_index += 1

        for col in range(len(in_row)):
            cell = out_ws.cell(row=out_index, column=col + 1, value=in_row[col])

            if index > 0:
                if col in col_format:
                    # handle special column formats
                    cell.number_format = col_format[col]

                if col in col_fills:
                    col_fills[col](cell)

            if col in col_alignment:
                cell.alignment = col_alignment[col]

    # do some ws dependent table fixup after the copy
    if 'post_fixup' in params:
        params['post_fixup'](in_ws, out_ws, title_dict)


    # now make a table of the data
    start_col = 'A'
    end_col = openpyxl.utils.cell.get_column_letter(len(title_dict))
    table_ref = f"{start_col}{out_starting_row}:{end_col}{out_starting_row + num_rows -1}"
    log.debug(f"table_ref '{table_ref}'")
    table = openpyxl.worksheet.table.Table(displayName=params['table_name'], ref=table_ref)
    out_ws.add_table(table)


    default_sheet_name = 'Sheet'
    if default_sheet_name in out_wb:
        del out_wb[default_sheet_name]

    out_wb.save(params['out_file_name'])


def process_title_row(row):
    """ process an xlrd title row, returning a map of column names to column indexes (origin zero) """
    name_dict = {}
    for col in range(len(row)):
        value = row[col]
        if value is not None and value != '':
            name_dict[value] = col

    return name_dict


def read_arrival_roster(session, config, firsttime):


    url = "https://volunteerconnection.redcross.org/"
    params = {
            'query_id': '1537756',
            'nd': 'clearreports_launch_admin',
            'reference': 'disaster',
            'prompt1': '1694',
            'prompt2': 'Arrival Date',
            'prompt3': 'No',
            'prompt4': 'Yes',
            'prompt5': '-1__,__',
            'output_format': 'xls',
            'run': 'Run',
            }
    headers = {
            #'accept': 'application/json, text/javascript, */*; q=0.01',
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9',
            'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'DNT': '1',
            'Host': 'volunteerconnection.redcross.org',
            'Origin': 'https://volunteerconnection.redcross.org',
            'Referer': 'https://volunteerconnection.redcross.org/',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'sec-gpc': '1',
            'User-Agent': '.Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
            }
    response = session.post(url, data=params, headers=headers, timeout=config.WEB_TIMEOUT)
    response.raise_for_status()


    # now poll for the report to be done
    params = {
            'nd': 'clearreports_auth',
            'init': 'xls',
            'query_id': '1537756',
            'prompt0': '1694',
            'prompt1': 'Arrival Date',
            'prompt2': 'No',
            'prompt3': 'Yes',
            'prompt4': "['-1']",
            }

    response = session.post(url, data=params, headers=headers, timeout=config.WEB_TIMEOUT)
    response.raise_for_status()

    log.debug(f"response.content { response.content }")
    anchor = response.html.find('a', first=True)
    href = anchor.attrs['href']

    url2 = url + href

    log.debug(f"url2 { url2 }")
    response = session.get(url2, timeout=config.WEB_TIMEOUT)
    response.raise_for_status()

    log.debug(f"retrieved document.  size is { len(response.text) }, type is '{ response.headers.get('content-type') }'")
    #log.debug(f"document: { response.content }")

    return response.content




def init_config():
    class AttrDict(dict):
        def __init__(self, *args, **kwargs):
            super(AttrDict, self).__init__(*args, **kwargs)
            self.__dict__ = self


    config_dotenv = dotenv.dotenv_values(verbose=True)

    config = AttrDict()
    for item in dir(config_static):
        if not item.startswith('__'):
            config[item] = getattr(config_static, item)

    #log.debug(f"config after copy: { config.keys() }")

    for key, val in config_dotenv.items():
        config[key] = val

    return config

def parse_args():
    parser = argparse.ArgumentParser(
            description="process support for the regional bootcamp mission card system",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")

    args = parser.parse_args()
    return args



if __name__ == "__main__":
    main()

