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

PARSER = argparse.ArgumentParser(description="Transforms a spreadsheet to RDF")
PARSER.add_argument("-s", "--spreadsheet", help="Google Spreadsheet")
PARSER.add_argument("-w", "--worksheet", help="Worksheet")
ARGS = PARSER.parse_args()

if ARGS.spreadsheet:
    G_SPREADSHEET_NAME = ARGS.spreadsheet
if ARGS.worksheet:
    G_WORKSHEET_NAME = ARGS.worksheet

LAST_UPDATE_CELL = "X2"
LAST_UPDATE_FORMAT = "%m/%d/%Y %H:%M:%S"
LAST_UPDATE_FILE = G_SPREADSHEET_NAME + "_last_update.txt"
CSV_FILE = G_SPREADSHEET_NAME + ".csv"
ENCODE = "utf-8"
TIME = str(datetime.now())
LAST_PROJECT_FILE = G_SPREADSHEET_NAME + "_last_project.txt"

REFINE_PROJECT_NAME = G_SPREADSHEET_NAME + "_" + TIME
REFINE_APPLY_JSON = "apply.json"
REFINE_RDF_FILE = G_SPREADSHEET_NAME + ".rdf"


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


def apply_operations(project, file_path, wait=True):
    """Transforms the OpenRefine project to RDF applying the JSON file"""
    json_data = open(file_path).read()
    resp_json = project.do_json('apply-operations', {'operations': json_data})
    if resp_json['code'] == 'pending' and wait:
        project.wait_until_idle()
        return 'ok'
    return resp_json['code']  # can be 'ok' or 'pending'


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

LOGGER = logging.getLogger("sp_log")
FORMAT = '%(asctime)s | %(message)s'
logging.basicConfig(format=FORMAT)
LOGGER.setLevel(logging.DEBUG)

SCOPE = ['https://spreadsheets.google.com/feeds']
CRED = ServiceAccountCredentials.from_json_keyfile_name(G_AUTH_JSON, SCOPE)
GC = gspread.authorize(CRED)

# Open the worksheet from spreadsheet
WKS = GC.open(G_SPREADSHEET_NAME).worksheet(G_WORKSHEET_NAME)

# Get the last updated time from the cell to compare with the one in the file
LASTDATE = WKS.acell(LAST_UPDATE_CELL).value
DATE_OBJ = datetime.strptime(LASTDATE, LAST_UPDATE_FORMAT)

FILE_EXISTS = os.path.isfile(LAST_UPDATE_FILE)

if FILE_EXISTS:
    LAST_DATE_2 = open(LAST_UPDATE_FILE).readline().rstrip()
    DATE_OBJ_2 = datetime.strptime(LAST_DATE_2, LAST_UPDATE_FORMAT)
    if DATE_OBJ == DATE_OBJ_2:
        LOGGER.debug("The spreadsheet is updated on date %s", DATE_OBJ_2)

if (not FILE_EXISTS) or (DATE_OBJ > DATE_OBJ_2):
    if not FILE_EXISTS:
        LOGGER.debug("Generating spreadsheet for date %s...", DATE_OBJ)
    elif DATE_OBJ > DATE_OBJ_2:
        LOGGER.debug("Updating spreadsheet to date %s...", DATE_OBJ)
    LAST_DATE_FILE = open(LAST_UPDATE_FILE, 'w')
    LAST_DATE_FILE.write("%s\n" % LASTDATE)
    LAST_DATE_FILE.close()

    LOGGER.debug("Exporting spreadsheet in CSV...")
    EXP_DATA = WKS.export(format='csv')
    EXP_VALUES = [(line.decode(ENCODE)) for line in EXP_DATA.splitlines()]
    EXP_FILE = open(CSV_FILE, 'w')
    for item in EXP_VALUES:
        EXP_FILE.write("%s\n" % item.encode(ENCODE))
    EXP_FILE.close()
    LOGGER.debug("Spreadsheet exported to file %s", CSV_FILE)

    LOGGER.debug("Opening connection with OpenRefine...")
    SERVER = refine.RefineServer()
    REFINE_INSTANCE = refine.Refine(SERVER)
    LOGGER.debug("Connected to OpenRefine")

    LOGGER.debug("Uploading CSV file to OpenRefine...")
    OPTIONS_JSON = get_options(REFINE_PROJECT_NAME, CSV_FILE)
    OPTS = {}
    NEW_STYLE_OPTIONS = dict(OPTS, **{
        'encoding': ENCODE,
    })
    PARAMS = {
        'options': json.dumps(NEW_STYLE_OPTIONS),
    }
    RESP = SERVER.urlopen('create-project-from-upload', OPTIONS_JSON, PARAMS)
    URL_PARAMS = urlparse.parse_qs(urlparse.urlparse(RESP.geturl()).query)

    if 'project' in URL_PARAMS:
        PROJECT_ID = URL_PARAMS['project'][0]
        LOGGER.debug("Created project with project id: %s", PROJECT_ID)
        PROJ = refine.RefineProject(PROJECT_ID)
        update_project_file(PROJECT_ID)
    else:
        raise Exception('Project not created')

    LOGGER.debug("Exporting project to RDF...")
    apply_operations(PROJ, REFINE_APPLY_JSON)
    export_project(PROJ, REFINE_RDF_FILE)
    if REFINE_RDF_FILE:
        LOGGER.debug("RDF exported to %s", REFINE_RDF_FILE)
