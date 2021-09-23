# Excel2SQL
Convert an Excel file to CSV and SQL scripts.

This project is originally from the article posted on https://www.oracle.com/technetwork/articles/saternos-tables-090560.html. The author is Casimir Saternos.

# Functionality

Reads an Excel Spread sheet (file name passed in as an argument).  

1. Creates a comma delimited (CSV) file for each sheet.

2. Creates a SQL script to create tables.
* For an Oracle database use external table definitions that reference the CSV files on the file system.
* For an PostgresQL database use table definitions and add a load.sql script that loads the CSV file into the table.

# Help

```
$ java -jar <path to>/Excel2SQL-<version>.jar --help
```

gives:

```
Usage: TableGenerator [options] spreadsheet...
  Options:
    --add-metadata
      Add metadata like sheet name and row number (starting from 1) in the 
      CSV? 
      Default: false
    --column-separator
      The column separator
      Default: ,
    --enclosure-string
      The enclosure string
      Default: "
    --encoding
      The encoding to use (default "windows-1252")
      Default: windows-1252
    --header-row-from
      The first header row (0 means no header)
      Default: 1
    --header-row-till
      The last header row (0 means no header)
      Default: 1
    --help, -h
      This help
    --one-table
      Each sheet to one table?
      Default: false
    --sheet-name-expression
      Sheet name(s) must match this regular expression
      Default: .*
    --sql-database
      The SQL database (Oracle, PostgresQL)
      Default: Oracle
    --sql-table-name
      A list of SQL table name(s) to use instead of the sheet name(s)
      Default: []
    --verbose
      Level of verbosity
      Default: 1
    --write-bom
      Write the BOM at the beginning of the file.
      Default: false
```

# Algorithm

There is just one workbook. Previous versions allowed more than one workbook but that got too difficult to maintain.

## Per worksheet

The sheet table name will be the worksheet name except when --one-table is used or --sql-table-name is used.

### For every row construct a sparse array of its cell values

In Java a sparse array is implemented by a HashMap.

Now:
1. If a cell is part of a header row add the header name (cell contents) to the previous contents of the corresponding header array element (with a space as separator), if any.
2. Else:
   - If there is no header and the cell is beyond the largest sheet column found till now, add the corresponding column name (A, B, ...) as the header name to the (sparse) header array and the cell itself to the (sparse) data array.
   - Else, it is just a data cell so add it to the (sparse) data array.

Every time a data cell is processed, the header name (via the column index) will be used to:
- either add the header name as a table column OR
- retrieve the column with that name

Now the cell value will be used to update the table column data type.

At the end of a data row, print the values to a CSV file (&lt;sheet&gt;.csv) taking care of the holes in the sparse data array.

# Enhancements

I have made the following enhancements:
* It is now a Maven project.
* Microsoft Excel Open XML Spreadsheet is now supported (xlsx extension).
* Dates are now supported.
* In the article the following is stated: "The second row in each spreadsheet is used to determine the length of a VARCHAR2 field. This row is also not included in the CSV datafile. ". This requirement has been removed. The Excel should just contain a header row and data. The type of a column (STRING, NUMERIC or DATE) is determined while processing the Excel and the maximum length for a STRING column field too.
* The newest Jakarta POI is used.
* All Excel cell types are supported, so FORMULA and BOOLEAN have been added, except for ERROR.
* The CSV output takes care of separators and double quotes in a cell. Such a cell is enclosed by double quotes and the double quote in the cell is duplicated. See also https://en.wikipedia.org/wiki/Comma-separated_values. The external table definition uses "FIELDS CSV WITHOUT EMBEDDED RECORD TERMINATORS" to support this.
* The CSV output file also contains the header, so it is just a copy of the worksheet. The external table skips that row by using "FIELD NAMES ALL FILES IGNORE".
* The external table definition uses double quoted identifiers ("This is my column") instead of (This is my column) in order to suppress DDL errors.
* The external table does not create a log file since that grows and grows...
* The column names are derived from the header or will be A, B, ... if no header is specified.
