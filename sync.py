#!/usr/bin/env python

import os
import re
import io
import sys
import logging
import argparse
import datetime

import xlrd
import dotenv
import O365

import init_logging
import config as config_static
import roster


log = logging.getLogger(__name__)


DATESTAMP = datetime.datetime.now().strftime("%Y-%m-%d")

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = init_config()

    credentials = (config.CLIENT_ID, config.CLIENT_SECRET)

    scopes = [
            'https://graph.microsoft.com/Files.ReadWrite.All',
            'https://graph.microsoft.com/Mail.Send',
            'https://graph.microsoft.com/Mail.Send.Shared',
            'https://graph.microsoft.com/offline_access',
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/User.ReadBasic.All',
            #'https://microsoft.sharepoint-df.com/AllSites.Read',
            #'https://microsoft.sharepoint-df.com/MyFiles.Read',
            #'https://microsoft.sharepoint-df.com/MyFiles.Write',
            'https://graph.microsoft.com/Sites.ReadWrite.All',
            'https://graph.microsoft.com/offline_access',
            'basic',
            ]

    token_backend = O365.FileSystemTokenBackend(token_path='.', token_filename="my_token.txt")
    account = O365.Account(credentials, scopes=scopes, token_backend=token_backend)

    if not account.is_authenticated:
        account.authenticate()
        if not account.is_authenticated:
            log.fatal(f"Cannot authenticate account")
            sys.exit(1)

    # this section was just for debugging token refresh; no longer needed
    #else:
    #    log.info(f"account { account }")
    #    connection = account.con
    #    result = connection.refresh_token()
    #    log.info(f"refreshed existing token; result { result }")


    #send_test_message(account)

    test_storage(account, config.WORKFORCE_DRIVE_ID, config)

    #test_sharepoint(account, config.WORKFORCE_SITE_ID)


def test_sharepoint(account, site_id):

    sharepoint = account.sharepoint()
    site = sharepoint.get_site(site_id)

    log.debug(f"got site { site }")


def test_storage(account, drive_id, config):

    storage = account.storage()

    drive = storage.get_drive(drive_id)
    log.debug(f"drive { drive }")

    root = drive.get_root_folder()
    log.debug(f"root { root }")

    for item in root.get_child_folders():
        log.debug(f"root child item: { item }")

    neil_notes = drive.get_item_by_path('/Workforce/Neil Notes/Mail Merge Spreadsheets')
    log.debug(f"neil_notes { neil_notes }")

    xls_dict = {}
    xlsx_dict = {}
    xls_pattern = re.compile(r'staff_roster(.*).xls')
    xlsx_pattern = re.compile(r'staffing(.*).xlsx')

    for entry in neil_notes.get_items():
        name = entry.name
        log.debug(f"entry: { name } is_file { entry.is_file } modified { entry.modified } modified_by { entry.modified_by }")

        xls_result = xls_pattern.fullmatch(name)
        xlsx_result = xlsx_pattern.fullmatch(name)
        if xls_result is not None:
            xls_dict[name] = entry
        elif xlsx_result is not None:
            xlsx_dict[name] = entry

    for name, entry in xls_dict.items():
        xls_result = xls_pattern.fullmatch(name)

        xlsx_name = xls_result.expand(r'staffing\1.xlsx')

        log.debug(f"pass 2: name { name } xlsx_name { xlsx_name }")

        # algorithm: we are going to process:
        #   any xls without an xlsx
        #   any xls that is newer than xlsx

        date_format = '%Y-%m-%d %H:%M:%S'
        if xlsx_name not in xlsx_dict:
            # no xlsx entry: process it
            log.info(f"converting { name } because no xlsx matching entry")
            convert(neil_notes, entry, xlsx_name, config)
        else:
            xlsx_entry = xlsx_dict[xlsx_name]
            if entry.modified > xlsx_entry.modified:
                log.info(f"converting { name }/{ entry.modified.strftime(date_format) } because xlsx { xlsx_name }/{xlsx_entry.modified.strftime(date_format) } is older.")
                convert(entry, xlsx_name)
            else:
                log.debug(f"ignoring { name }/{ entry.modified.strftime(date_format) } because xlsx { xlsx_name }/{xlsx_entry.modified.strftime(date_format) } is newer.")



def convert(folder, entry, new_name, config):
    input_io = io.BytesIO()
    entry.download(output=input_io)

    input_buffer = input_io.getvalue()
    input_io.close()

    input_wb = xlrd.open_workbook(file_contents=input_buffer)
    input_buffer = None

    output_wb = roster.make_workbook(input_wb, config)

    output_io = io.BytesIO()
    output_wb.save(output_io)
    output_buffer = output_io.getvalue()
    output_io.close()

    stream_io = io.BytesIO(output_buffer)
    stream_len = len(output_buffer)
    output_buffer = None

    new_file = folder.upload_file(None, stream=stream_io, stream_size=stream_len, item_name=new_name)

    log.info(f"uploaded { new_file.name }: modified { new_file.modified } by { new_file.modified_by }")



def send_test_message(account):
    m = account.new_message()
    m.to.add('generic@askneil.com')
    m.subject = 'Testing'
    m.body = 'Test message using python o365 package'
    m.send()

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

