# spreadsheet_to_rdf

This script is used to download automatically a Google Spreadsheet as CSV and transform it in RDF via OpenRefine.

The script is based on Python 2.7.10 and it requires 2 libraries to work:
 * gspread (https://github.com/burnash/gspread), to retrieve data from Google Spreadsheet
 * refine-client (https://github.com/PaulMakepeace/refine-client-py), to send the CSV to OpenRefine and retrieve the RDF

In order to retrieve data from Google Spreadsheet, you need to:
 * get a JSON file from  Google Developer Console, see: http://gspread.readthedocs.io/en/latest/oauth2.html, the JSON file is used by the script in the variable G_AUTH_JSON
 * create a Spreadsheet and fill a worksheet, these are used by the script in the variables G_SPREADSHEET_NAME and G_WORKSHEET_NAME
 * share your Google Spreadsheet to the email contained in your JSON file
 * put in your Google Spreadsheet a cell which contains the last modified date with the following script (from Tools -> Script editor...):

```
function onEdit() {
  var s = SpreadsheetApp.getActiveSheet();
 if( s.getName() == "CPSV-AP classes and properties" ) { //checks that we're on the correct sheet
  var r = s.getActiveCell();
  if( r.getColumn() != 2 ) { //checks the column
    var row = r.getRow();
    var time = new Date();
    time = Utilities.formatDate(time, "GMT+02:00", "MM/dd/yyyy HH:mm:ss");
    SpreadsheetApp.getActiveSheet().getRange('X2').setValue(time);
  };
 };
}
```

The cell is used by the script in the variables: LAST_UPDATE_CELL and LAST_UPDATE_FORMAT
The cell value is stored in a file which is checked against the value of the cell, if the cell is more updated than the file, or the file does not exist (first time execution) the CSV is downloaded and extracted.

In order to contact OpenRefine, you need to:
 * Download OpenRefine, I have version 2.6 RC2 but the refine-client has been tested on 2.5, so it should work on 2.5
 * Download the RDF extension zip file from: http://openrefine.org/download.html#list-of-extensions and copy it in your extensions folder
 * set environment variables to the address and ports, in my case it is running on port 4000 so I set the environment variable OPENREFINE_PORT to 4000
 * save the JSON file from the OpenRefine that will perform the transformation (it means that you have to do once manually the transformation), the JSON file is used by the script in the variable REFINE_APPLY_JSON
 
 