# Excel2SQL
Convert an Excel file to a CSV file per sheet and SQL script(s) to load data into a table using the CSV file.

This project is originally from the article posted on https://www.oracle.com/technetwork/articles/saternos-tables-090560.html. The author is Casimir Saternos.

# Functionality

Reads an Excel Spread work book (file name passed in as an argument).  

1. Creates a comma delimited (CSV) file for each work sheet.

2. Creates SQL scripts:
   * table.sql: to create a table for each work sheet.
     For an Oracle database use external table definitions that reference the CSV files on the file system.
     For an PostgresQL database use table definitions.
   * delete.sql (only for an PostgresQL database):
     to delete work sheet data before loading.
   * load.sql (only for an PostgresQL database):
     load the CSV file into the table(s).

# Help

```
$ mvn clean package
$ java -jar <path to>/Excel2SQL-<version>.jar --help
```

gives:

```
Usage: TableGenerator [options] spreadsheet
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
    --sheet-name-expression
      Sheet name(s) must match this regular expression
      Default: .*
    --sql-database
      The SQL database (Oracle, PostgresQL)
      Default: Oracle
    --sql-table-names
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

There is just one spreadsheet (work book) to process.

## Per work sheet

The sheet table name will be the work sheet name except when --sql-table-names
is used. In that case the --sql-table-names command line option must have
exactly the same number of elements as there are work sheets.

### For every row construct a sparse array of its cell values

Jakarta POI is used to read the work book and sheets. Only rows and columns
with contents are returned by the iterators. So you may start with Excel row 5
and within that column 4 (D).  So you need to store the data as a sparse
array. Every time a cell is processed the data type and its maximum length
may be updated. Since it is possible that cell A1 has a string length of 10
but D1 a length of 20, you know the maximum length only at the end of the
sheet. And while you are processing a column, the data type may change too. If
there is more than one data type detected the final data type will be string.

Another problem is that you do not know the exact number of columns after the first
row. It is possible that POI returns cells A2 and A5 and later on D4 and D8,
etcetera.  So you can not print the row in between: you have to wait till the
end of a sheet to know the number of columns because you have to print the
exact number of column separators (final column count minus 1).

This is algorithm for processing a cell:
1. If a cell is part of a header row add the header name (cell contents)
   to the previous contents of the corresponding header array element (with a space as separator), if any.
2. Else:
   * If there is no header and the cell is beyond the largest sheet column found till now,
     add the corresponding column name (A, B, ...) as the header name
     to the (sparse) header array and the cell itself to the (sparse) data array
   * Else, it is just a data cell so add it to the (sparse) data array.

Every time a data cell is processed, the header name (via the column index) will be used to:
* either add the header name as a table column OR
* retrieve the column with that name

Now the cell value will be used to update the table column data type.

At the end of a data row, print the values to a CSV file (&lt;sheet&gt;.csv) taking care of the holes in the sparse data array.

# Enhancements

I have made the following enhancements to the original code from Casimir Saternos:
* It is now a Maven project.
* Microsoft Excel Open XML Spreadsheet is now supported (xlsx extension).
* Dates are now supported.
* In the article the following is stated: "The second row in each spreadsheet is used to determine the length of a VARCHAR2 field. This row is also not included in the CSV datafile. ". This requirement has been removed. The Excel should just contain a header row and data. The type of a column (STRING, NUMERIC or DATE) is determined while processing the Excel and the maximum length for a STRING column field too.
* The newest Jakarta POI is used.
* All Excel cell types are supported, so FORMULA and BOOLEAN have been added, except for ERROR.
* The CSV output takes care of separators and double quotes in a cell. Such a cell is enclosed by double quotes and the double quote in the cell is duplicated. See also https://en.wikipedia.org/wiki/Comma-separated_values. The external table definition uses "FIELDS CSV WITHOUT EMBEDDED RECORD TERMINATORS" to support this.
* You can have a header consisting of 0 (so no header), 1 (the standard) or more rows. Multiple row header columns are concatened into one header column separated by spaces.
* You can specify on which row the header starts or ends.
* The column names are derived from the header or will be A, B, ... if no header is specified.
* The CSV output file also contains the header, so it is just a copy of the worksheet. The external table skips that row by using "FIELD NAMES ALL FILES IGNORE".
* The external table definition uses double quoted identifiers ("This is my column") instead of (This is my column) in order to suppress DDL errors.
* The external table does not create a log file since that grows and grows...
* PostgreSQL as a database option has been added.
* Checkstyle has been added.
