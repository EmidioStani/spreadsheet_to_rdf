"""
.. module:: spreadsheet_to_rdf
   :platform: Unix, Windows
   :synopsis: It transforms a Google spreadsheet into in RDF via OpenRefine
.. moduleauthor:: Emidio Stani <emidio.stani@be.pwc.com>
"""

__author__ = 'Emidio Stani, PwC EU Services'

import gspread
import sys
import os
import json
import urlparse
import logging
import argparse

from oauth2client.service_account import ServiceAccountCredentials
from google.refine import refine
from datetime import datetime

G_AUTH_JSON = "cpsv-ap-labels-a488e0f3d7bb.json"
G_SPREADSHEET_NAME = "CPSV-AP Multilingual labels"
G_WORKSHEET_NAME = "CPSV-AP classes and properties"
G_UPDATE_CELL = "None"
APPLY_JSON = "apply.json"

PARSER = argparse.ArgumentParser(description="Transforms a spreadsheet to RDF")
PARSER.add_argument("-a", "--auth", help="JSON file for authentication")
PARSER.add_argument("-s", "--spreadsheet", help="Google Spreadsheet name")
PARSER.add_argument("-w", "--worksheet", help="Worksheet name")
PARSER.add_argument("-c", "--cell", help="Cell containing last update date")
PARSER.add_argument("-t", "--transform", help="JSON file for transformation")
ARGS = PARSER.parse_args()

if ARGS.auth:
    G_AUTH_JSON = ARGS.auth
if ARGS.spreadsheet:
    G_SPREADSHEET_NAME = ARGS.spreadsheet
if ARGS.worksheet:
    G_WORKSHEET_NAME = ARGS.worksheet
if ARGS.cell:
    G_UPDATE_CELL = ARGS.cell
if ARGS.transform:
    APPLY_JSON = ARGS.transform

G_UPDATE_FORMAT = "%m/%d/%Y %H:%M:%S"
ENCODE = "utf-8"
UPDATE_FILE = G_SPREADSHEET_NAME + "_last_update.txt"
CSV_FILE = G_SPREADSHEET_NAME + ".csv"
TIME = str(datetime.now())
LAST_PROJECT_FILE = G_SPREADSHEET_NAME + "_last_project.txt"
PROJECT_NAME = G_SPREADSHEET_NAME + "_" + TIME
RDF_FILE = G_SPREADSHEET_NAME + ".rdf"


def check_last_update_cell(last_update_cell, date_format, output_file):
    """Check if CSV need to be exported depending on the date in the cell"""
    if last_update_cell != "None":
        LOGGER.debug("Spreadsheet checked against cell: %s", last_update_cell)
        lastdate = WKS.acell(last_update_cell).value
        date_obj = datetime.strptime(lastdate, date_format)
        file_exists = os.path.isfile(output_file)

        if file_exists:
            last_date_2 = open(output_file).readline().rstrip()
            date_obj_2 = datetime.strptime(last_date_2, date_format)
            if date_obj == date_obj_2:
                LOGGER.debug("The spreadsheet is updated on: %s", date_obj_2)
                return False

        if (not file_exists) or (date_obj > date_obj_2):
            if not file_exists:
                LOGGER.debug("Generating spreadsheet for date: %s", date_obj)
            elif date_obj > date_obj_2:
                LOGGER.debug("Updating spreadsheet to date: %s", date_obj)
            last_date_file = open(output_file, 'w')
            last_date_file.write("%s\n" % lastdate)
            last_date_file.close()
            return True
    else:
        return True


def export_spreadsheet_to_csv(worksheet, encoding, output_file):
    """Export Google Spreadsheet to CSV file"""
    exp_data = worksheet.export(format='csv')
    exp_values = [(line.decode(encoding)) for line in exp_data.splitlines()]
    exp_file = open(output_file, 'w')
    for item in exp_values:
        exp_file.write("%s\n" % item.encode(encoding))
    exp_file.close()
    LOGGER.debug("Spreadsheet exported to file: %s", output_file)


def get_options(project_name, csv_file):
    """Prepares the JSON file for OpenRefine with CSV file and project name"""
    options = {
        'format': 'text/line-based/*sv',
        'encoding': '',
        'separator': ',',
        'ignore-lines': '-1',
        'header-lines': '0',
        'skip-data-lines': '0',
        'limit': '-1',
        'store-blank-rows': 'true',
        'guess-cell-value-types': 'true',
        'process-quotes': 'true',
        'store-blank-cells-as-nulls': 'true',
        'include-file-sources': 'false',
    }
    options['project-file'] = {
        'fd': open(csv_file),
        'filename': csv_file,
    }

    options['project-name'] = project_name

    return options


def update_project_file(project_id):
    """Write on file the project id."""
    if not os.path.exists(LAST_PROJECT_FILE):
        update_file = open(LAST_PROJECT_FILE, 'w')
        update_file.write("%s\n" % project_id)
        update_file.close()
    else:
        update_file = open(LAST_PROJECT_FILE, 'r')
        last_project_created = update_file.readline().rstrip()
        update_file.close()
        LOGGER.debug("Deleting project id: %s", last_project_created)
        refine.RefineProject(last_project_created).delete()
        update_file = open(LAST_PROJECT_FILE, 'w')
        update_file.write("%s\n" % project_id)
        update_file.close()


def apply_operations(project, file_path, wait=True):
    """Transforms the OpenRefine project to RDF applying the JSON file"""
    json_data = open(file_path).read()
    resp_json = project.do_json('apply-operations', {'operations': json_data})
    if resp_json['code'] == 'pending' and wait:
        project.wait_until_idle()
        return 'ok'
    return resp_json['code']  # can be 'ok' or 'pending'


def export_project(project, output=False):
    """Dump a project to stdout or output file."""
    export_format = 'rdf'
    if output:
        ext = os.path.splitext(output)[1][1:]     # 'xls'
        if ext:
            export_format = ext.lower()
        output = open(output, 'wb')
    else:
        output = sys.stdout
    output.writelines(project.export(export_format=export_format))
    output.close()


def export_csv_to_rdf(proj_name, input_file, encoding, json_file, output_file):
    """Exports CSV to RDF via OpenRefine using a JSON file."""
    server = refine.RefineServer()
    LOGGER.debug("Connected to OpenRefine")

    options_json = get_options(proj_name, input_file)
    opts = {}
    new_style_options = dict(opts, **{
        'encoding': encoding,
    })
    params = {
        'options': json.dumps(new_style_options),
    }
    resp = server.urlopen('create-project-from-upload', options_json, params)
    url_params = urlparse.parse_qs(urlparse.urlparse(resp.geturl()).query)

    if 'project' in url_params:
        project_id = url_params['project'][0]
        LOGGER.debug("Created project with project id: %s", project_id)
        proj = refine.RefineProject(project_id)
        update_project_file(project_id)
    else:
        raise Exception('Project not created')

    apply_operations(proj, json_file)
    export_project(proj, output_file)
    if output_file:
        LOGGER.debug("RDF exported to: %s", output_file)

LOGGER = logging.getLogger("sp_log")
FORMAT = '%(asctime)s | %(message)s'
logging.basicConfig(format=FORMAT)
LOGGER.setLevel(logging.DEBUG)

SCOPE = ['https://spreadsheets.google.com/feeds']
CRED = ServiceAccountCredentials.from_json_keyfile_name(G_AUTH_JSON, SCOPE)
GC = gspread.authorize(CRED)
LOGGER.debug("Authenticated on Google with: %s", G_AUTH_JSON)

SDS = GC.open(G_SPREADSHEET_NAME)
LOGGER.debug("Opened spreadsheet: %s", G_SPREADSHEET_NAME)
WKS = SDS.worksheet(G_WORKSHEET_NAME)
LOGGER.debug("Opened worksheet: %s", G_WORKSHEET_NAME)

CHECK = check_last_update_cell(G_UPDATE_CELL, G_UPDATE_FORMAT, UPDATE_FILE)
if CHECK:
    export_spreadsheet_to_csv(WKS, ENCODE, CSV_FILE)
    export_csv_to_rdf(PROJECT_NAME, CSV_FILE, ENCODE, APPLY_JSON, RDF_FILE)
