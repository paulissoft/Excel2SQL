# Excel2SQL
Convert an Excel file to CSV and SQL scripts.

This project is originally from the article posted on https://www.oracle.com/technetwork/articles/saternos-tables-090560.html. The author is Casimir Saternos.

# Functionality

Reads an Excel Spread sheet (file name passed in as an argument).  

1. Creates a comma delimitted (.csv) file for each sheet.

2. Creates a SQL script to create tables.
* For an Oracle database use external table definitions that reference the .csv files on the file system.
* For an PostgresQL database use table definitions and add a load.sql script that loads the CSV file into the table.

# Enhancements

I have made the following enhancements:
* It is now a Maven project.
* Microsoft Excel Open XML Spreadsheet is now supported (xlsx extension).
* Dates are now supported.
* In the article the following is stated: "The second row in each spreadsheet is used to determine the length of a VARCHAR2 field. This row is also not included in the .csv datafile. ". This requirement has been removed. The Excel should just contain a header row and data. The type of a column (STRING, NUMERIC or DATE) is determined while processing the Excel and the maximum length for a STRING column field too.
* The newest Jakarta POI is used.
* All Excel cell types are supported, so FORMULA and BOOLEAN have been added, except for ERROR.
* The CSV output takes care of separators and double quotes in a cell. Such a cell is enclosed by double quotes and the double quote in the cell is duplicated. See also https://en.wikipedia.org/wiki/Comma-separated_values. The external table definition uses "FIELDS CSV WITHOUT EMBEDDED RECORD TERMINATORS" to support this.
* The CSV output file also contains the header, so it is just a copy of the worksheet. The external table skips that row by using "FIELD NAMES ALL FILES IGNORE".
* The external table definition uses double quoted identifiers ("This is my column") instead of (This is my column) in order to suppress DDL errors.
* The external table does not create a log file since that grows and grows...
* The column names are derived from the header or will be A, B, ... if no header is specified.
