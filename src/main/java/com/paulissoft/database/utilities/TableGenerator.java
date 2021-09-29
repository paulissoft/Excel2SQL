package com.paulissoft.database.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.TreeSet;
import java.util.SortedSet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook; // interface
import org.apache.poi.ss.usermodel.Sheet; // interface
import org.apache.poi.ss.usermodel.Row; // interface
import org.apache.poi.ss.usermodel.Cell; // interface
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.beust.jcommander.JCommander;

/**
 * This program is run at the command line. Given one
 * spreadsheet, a script is generated to create the table(s)
 * that references the data from the spreadsheet.
 *
 * @author Casimir Saternos
 * @version 1.0
 */
public class TableGenerator {
    /** The run-time settings. */
    private static Settings settings = new Settings();

    private void info(final String str) {
        if (settings.verbose >= 1) {
            System.out.println("INFO: " + str);
        }
    }

    private void debug(final String str) {
        if (settings.verbose >= 2) {
            System.out.println("DEBUG: " + str);
        }
    }

    /** The data formatter. */
    private DataFormatter dataFormatter = new DataFormatter();

    /** The formula evaluator. */
    private FormulaEvaluator formulaEvaluator = null;

    /**
     * The main procedure.
     *
     * @param args  The command line arguments
     */
    public static void main(final String... args) throws java.io.IOException {
        TableGenerator generator = new TableGenerator();

        JCommander jc = JCommander.newBuilder()
            .addObject(settings)
            .build();

        jc.setProgramName("TableGenerator");

        try {
            jc.parse(args);

            if (settings.help) {
                jc.usage();
            } else {
                // Check file exists as a regular file
                File f = new File(settings.spreadsheet);

                try {
                    assert (f.exists() && !f.isDirectory());
                } catch (AssertionError e) {
                    System.err.println("File '" + settings.spreadsheet + "' does not exist or is not a regular file");
                    throw e;
                }
            }
        } catch (Exception e) {
            jc.usage();
            throw e;
        }

        if (!settings.help) {
            generator.execute();
        }
    }

    // All sheets in the workbook use the following constants.

    /**
     * List of table definitions.
     */
    private HashSet<String> tables; // test unique table names

    /** Tables.sql file descriptor. */
    private PrintStream tablesSql = null;

    /** delete.sql file descriptor. */
    private PrintStream deleteSql = null;

    /** load.sql file descriptor. */
    private PrintStream loadSql = null;

    /** The constructor. */
    public TableGenerator() {
        this.tables = new HashSet<String>();
    }

    /**
     *  Open the specified .xls or .xlsx and process it.
     */
    private void execute() throws java.io.IOException {
        info("Begin processing.");

        info("Using working directory " + new File(Settings.PWD).getAbsolutePath());

        tablesSql = open("tables.sql", "UTF-8", false, false);

        tablesSql.print(Table.preamble(settings));

        if (settings.sqlDatabase.equals(Settings.POSTGRESQL)) {
            deleteSql = open("delete.sql", "UTF-8", false, false);
            loadSql = open("load.sql", "UTF-8", false, false);
        }

        final String spreadsheet = (new File(settings.spreadsheet)).getAbsolutePath();

        try {
            Workbook wb;

            try {
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(spreadsheet));
                wb = new HSSFWorkbook(fs);
            } catch (OfficeXmlFileException e) {
                wb = new XSSFWorkbook(new FileInputStream(spreadsheet));
            }

            formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();

            formulaEvaluator.setIgnoreMissingWorkbooks(true);

            for (Sheet sheet : wb) {
                for (Row r : sheet) {
                    for (Cell c : r) {
                        switch (c.getCellType()) {
                        case FORMULA:
                            formulaEvaluator.evaluateFormulaCell(c);
                            break;

                        default:
                            break;
                        }
                    }
                }
            }

            info("Processing workbook " + spreadsheet);

            processWorkbook(wb);
        } catch (Exception e) {
            e.printStackTrace();

            throw e;
        }

        tablesSql.close();

        if (deleteSql != null) {
            deleteSql.close();
        }

        if (loadSql != null) {
            loadSql.close();
        }

        info("Processing complete.");
    }

    /**
     * Prcess a worksheet.
     *
     * @param sheet
     * @param table
     */
    private void processSheet(final Sheet sheet, final Table table) throws java.io.IOException {
        //Write out a .csv file based upon the sheet
        if (writeCsv(sheet, table)) {

            debug("Adding DDL for table " + table.getName());

            // Add the ddl for the table to the script
            tablesSql.print(table.getDdl());

            if (deleteSql != null) {
                if (settings.addMetadata) {
                    deleteSql.print("delete from " + table.getName() + " where \"SHEET\" = '" + sheet.getSheetName() + "'");
                } else {
                    deleteSql.print("truncate table " + table.getName());
                }
                deleteSql.print(";" + Settings.NL);
            }

            if (loadSql != null) {
                loadSql.print("\\copy " + table.getName() + "(");
                for (int c = 0; c < table.getNrColumns(); c++) {
                    loadSql.print((c > 0 ? ", " : "") + table.getColumn(c).getName());
                }
                loadSql.print(") from '" + table.getLocation() + "' with ( format CSV );" + Settings.NL);
            }
        }
    }

    /**
     * Iterate through each sheet in the workbook and process it.
     *
     * @param wb
     */
    private void processWorkbook(final Workbook wb) throws java.io.IOException {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);

            if (!wb.getSheetName(i).matches(settings.sheetNameExpression)) {
                info("Skipping sheet " + i + " (" + wb.getSheetName(i)
                     + ") since it does not match \"" + settings.sheetNameExpression + "\"");
                continue;
            } else {
                info("Processing sheet " + i + " (" + wb.getSheetName(i) + ")");
            }

            Table table = null;

            final String tableName = (i < settings.tableNames.size() ? settings.tableNames.get(i) : wb.getSheetName(i));

            table = new Table(tableName, settings);
            if (!tables.add(table.getName())) {
                throw new RuntimeException("Table " + table.getName() + " is not unique");
            }

            processSheet(sheet, table);

            info("Table " + table.getName() + " processed.");
        }
    }

    private String getStringValue(final Cell cell, final TableColumn col) {
        // final String value = cell.getStringCellValue();
        final String value = cell.getRichStringCellValue().getString();

        col.setStringLength(value.length());

        if (value != null && value.length() > 0) {
            debug("getStringValue for column '" + col.getName()
                  + "', value '" + value + "'"
                  + "', length " + col.getStringLength()
                  + " and format " + cell.getCellStyle().getDataFormatString());
        }

        return value;
    }

    private String getNumericValue(final Cell cell, final TableColumn col) {
        String value = null;

        // Test if a date! See https://poi.apache.org/help/faq.html
        final Boolean isDate = DateUtil.isCellDateFormatted(cell);

        if (isDate) {
            Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

            value = dateFormat.format(date);

            col.setDateLength(value.length());

            debug("getNumericValue for date column '" + col.getName()
                  + "', value '" + value
                  + "', length " + col.getDateLength()
                  + " and format " + cell.getCellStyle().getDataFormatString());
        } else {
            // Remove trailing % and remove thousands separator (comma)
            value = dataFormatter.formatCellValue(cell, formulaEvaluator).replace(",", "");

            if (value.endsWith("%")) {
                final int unit = 100;
                // replace the percent symbol and divide by 100
                value = Double.valueOf(Double.valueOf(value.replace("%", "")) / unit).toString();
            }

            String[] parts = value.split("\\.");

            for (int i = 0; i < parts.length; i++) {
                debug("parts[" + i + "]= '" + parts[i] + "'");
            }

            // store the length first since it may be important in setType()
            col.setNumericLength((parts.length == 1 ? parts[0].length() : parts[0].length() + parts[1].length()));
            col.setNumericPrecision((parts.length == 1 ? 0 : parts[1].length()));

            debug("getNumericValue for numeric column '" + col.getName()
                  + "', value '" + value
                  + "', length " + col.getNumericLength()
                  + ", precision " + col.getNumericPrecision()
                  + " and format " + cell.getCellStyle().getDataFormatString());
        }

        return value;
    }

    private String getBooleanValue(final Cell cell, final TableColumn col) {
        final String value = "" + cell.getBooleanCellValue();

        // Treat it as a STRING.
        col.setStringLength(value.length());

        debug("getBooleanValue for column '" + col.getName() + "' and value '" + value + "'");
        debug("cell: " + cell.toString());

        return value;
    }

    /**
     * Open a file on the file system.
     *
     * @param filename
     * @param encoding
     * @param writeBOM
     * @param append
     *
     * @return The file descriptor
     */
    private PrintStream open(final String filename,
                             final String encoding,
                             final boolean writeBOM,
                             final boolean append) throws java.io.IOException {

        try {
            PrintStream out = new PrintStream(new FileOutputStream(filename, append), false, encoding);

            // write the BOM?
            if (!append && encoding.equals("UTF-8") && writeBOM) {
                final byte byte1 = (byte) 0xEF;
                final byte byte2 = (byte) 0xBB;
                final byte byte3 = (byte) 0xBF;
                final byte[] bom = {byte1, byte2, byte3};
                out.write(bom);
            }

            return out;
        } catch (Exception e) {
            e.printStackTrace();

            throw e;
        }
    }

    /**
     * Parse and write the CSV file.
     *
     * @param sheet
     * @param table
     *
     * @return Did this CSV have contents?
     */
    private Boolean writeCsv(final Sheet sheet, final Table table) throws java.io.IOException {
        HashMap<Integer, String> headerRow = new HashMap<Integer, String>();
        String progress = null;
        int nrNonEmptyRows = 0;
        Iterator<Row> rowIterator = sheet.rowIterator();
        PrintStream csvFile = null;
        // we need to process all rows before we defintively know many columns there are
        HashMap<Integer, HashMap<Integer, String>> dataRows = new HashMap<Integer, HashMap<Integer, String>>();

        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            HashMap<Integer, String> dataRow = new HashMap<Integer, String>();
            final boolean hasHeader = settings.headerRowFrom > 0;
            final boolean isHeaderRow =
                hasHeader && row.getRowNum() >= settings.headerRowFrom - 1 && row.getRowNum() <= settings.headerRowTill - 1;
            final boolean isDataRow = !hasHeader || row.getRowNum() > settings.headerRowTill - 1;
            final boolean lastDataRow = hasHeader && row.getRowNum() == settings.headerRowTill - 1;

            debug("row number          : " + row.getRowNum());
            debug("first cell number   : " + row.getFirstCellNum());
            debug("last cell number + 1: " + row.getLastCellNum());

            debug("Processing Excel row " + (row.getRowNum() + 1));

            final int displayAfter = 10;

            switch (nrNonEmptyRows++ % displayAfter) {
            case 0:
                progress = "Processing row " + nrNonEmptyRows;
                break;

            case displayAfter - 1:
                progress += ".";
                info(progress);
                progress = null;
                break;

            default:
                progress += ".";
            }

            while (cellIterator.hasNext()) {
                final Cell cell = cellIterator.next();

                debug("Processing Excel column " + (cell.getColumnIndex() + 1));
                debug("Number of table columns: " + table.getNrColumns());
                debug("cell address: " + cell.getAddress() + "; cell column index: " + cell.getColumnIndex());

                /*
                 * 1. If a cell is part of a header row add the header name (cell contents)
                      to the previous contents of the corresponding header array element (with a space as separator), if any
                 * 2. Else:
                 *    a) If there is no header and the cell is beyond the largest sheet column found till now,
                         add the corresponding column name (A, B, ...) as the header name
                         to the (sparse) header array and the cell itself to the (sparse) data array
                 *    b) Else, it is just a data cell so add it to the (sparse) data array.
                 *
                 * Every time a data cell is processed, the header name (via the column index) will be used to:
                 * - either add the header name as a table column OR
                 * - retrieve the column with that name
                 *
                 * Now the cell value will be used to update the table column data type.
                 *
                 * At the end of a data row, print the values to a CSV file (<sheet>.csv)
                 * taking care of the holes in the sparse data array.
                 *
                 */

                // Sometimes there may be cells missing so after cell column index 0 may come cell column index 2.
                // But not for the header!
                if (isHeaderRow) {
                    // See note 1 above.
                    String header = headerRow.get(cell.getColumnIndex());
                    String value;

                    // Some names are just numbers, strangely enough (column name 14)
                    try {
                        value = cell.getRichStringCellValue().getString();
                        // string?
                    } catch (IllegalStateException e1) {
                        // java.lang.IllegalStateException: Cannot get a STRING value from a NUMERIC cell
                        value = dataFormatter.formatCellValue(cell);
                    }

                    if (header != null) {
                        header += " " + value;
                    } else {
                        header = value;
                    }

                    headerRow.put(cell.getColumnIndex(), header); // add or replace header
                } else if (isDataRow) {
                    if (!hasHeader && (headerRow.isEmpty() || cell.getColumnIndex() > Collections.max(headerRow.keySet()))) {
                        // See note 2a (first part) above.
                        final Integer lastKey = (headerRow.isEmpty() ? -1 : Collections.max(headerRow.keySet()));

                        for (Integer key = lastKey + 1; key <= cell.getColumnIndex(); key++) {
                            final TableColumn col = new TableColumn(settings);
                            final String header = number2excelColumnName(key);

                            headerRow.put(key, header); // add or replace header
                            col.setName(header);
                            debug("setting header column " + (key + 1) + " to '" + header + "' (1)");
                            table.addColumn(col);
                        }
                        assert (cell.getColumnIndex() == Collections.max(headerRow.keySet()));
                    }

                    // Note 2a (second part) and 2b.
                    assert (cell.getColumnIndex() < table.getNrColumns());

                    // Add the value to the sparse data array.
                    final TableColumn col = table.getColumn(cell.getColumnIndex());
                    String value;

                    // column type can switch from string to numeric but not vice versa
                    switch (cell.getCellType()) {
                    case FORMULA:
                        // try in this order: numbers, strings (a boolean is a string in the database) and booleans
                        try {
                            value = getNumericValue(cell, col);
                        } catch (IllegalStateException e1) {
                            try {
                                value = getStringValue(cell, col);
                            } catch (IllegalStateException e2) {
                                value = getBooleanValue(cell, col);
                            }
                        }
                        break;

                    case BLANK:
                    case STRING:
                        value = getStringValue(cell, col);
                        break;

                    case BOOLEAN:
                        value = getBooleanValue(cell, col);
                        break;

                    case NUMERIC:
                        value = getNumericValue(cell, col);
                        break;

                    default:
                        throw new RuntimeException("Cell Type of cell " + col.getName() + " unknown: " + cell.getCellType());
                    }

                    // see https://en.wikipedia.org/wiki/Comma-separated_values
                    value.replace(table.getEnclosureString(), table.getEnclosureString() + table.getEnclosureString());
                    if (value.contains(table.getEnclosureString()) || value.contains(table.getFieldSeparator())) {
                        value = table.getEnclosureString() + value + table.getEnclosureString();
                    }
                    dataRow.put(cell.getColumnIndex(), value);
                }
            } // while (cellIterator.hasNext()) {

            // Create table columns
            if (lastDataRow) {
                // add table columns sorted on key
                final SortedSet<Integer> keys = new TreeSet<Integer>(headerRow.keySet());
                int lastKey = -1;

                for (Integer key : keys) {
                    TableColumn col;
                    String header;

                    // missing headers get the Excel column A, B, ...
                    while (++lastKey < key) {
                        col = new TableColumn(settings);
                        header = number2excelColumnName(lastKey);
                        col.setName(header);

                        debug("setting header column " + (lastKey + 1) + " to '" + header + "' (2)");
                        table.addColumn(col);
                    }

                    col = new TableColumn(settings);
                    header = headerRow.get(key);
                    col.setName(header);

                    debug("setting header column " + (key + 1) + " to '" + header + "' (3)");
                    table.addColumn(col);
                }
            }

            if (!isEmptyRow(dataRow)) {
                dataRows.put(row.getRowNum(), dataRow);
            } else {
                debug("Skipping row " + (row.getRowNum() + 1) + " since it is empty");
            }
        }

        if (progress != null) {
            info(progress);
        }

        System.out.println("");

        if (dataRows.size() > 0) {
            csvFile = open(table.getLocation(), settings.encoding, settings.writeBOM, false);

            final SortedSet<Integer> rows = new TreeSet<Integer>(dataRows.keySet());

            for (Integer row : rows) {
                final HashMap<Integer, String> dataRow = dataRows.get(row);

                final SortedSet<Integer> columns = new TreeSet<Integer>(dataRow.keySet());

                int lastColumn = 0;
                int nrSeparatorsToPrint = 0;
                int nrSeparatorsPrinted = 0;

                if (settings.addMetadata) {
                    csvFile.print(sheet.getSheetName() + table.getFieldSeparator());
                    csvFile.print(String.valueOf(row + 1) + table.getFieldSeparator());
                }

                for (Integer column : columns) {
                    final String col = dataRow.get(column);

                    // missing columns get a field separator
                    nrSeparatorsToPrint = column - lastColumn;
                    while (nrSeparatorsToPrint-- > 0) {
                        csvFile.print(table.getFieldSeparator());
                        nrSeparatorsPrinted++;
                    }

                    csvFile.print(col);
                    lastColumn = column;
                }

                // print separators after the last column if that was not the last table column
                nrSeparatorsToPrint = table.getNrColumns() - 1 - lastColumn;
                while (nrSeparatorsToPrint-- > 0) {
                    csvFile.print(table.getFieldSeparator());
                    nrSeparatorsPrinted++;
                }

                if (nrSeparatorsPrinted != table.getNrColumns() - 1) {
                    throw new RuntimeException("nrSeparatorsPrinted (" + nrSeparatorsPrinted + ") != "
                                               + "table.getNrColumns() - 1 (" + (table.getNrColumns() - 1) + ")");
                }

                csvFile.print(Settings.NL);
            }

            csvFile.close();
        } else {
            System.out.println("WARNING: Sheet does not contain data");
        }

        // >Add two table columns AFTER the CSV data has been printed
        if (settings.addMetadata) {
            TableColumn col = new TableColumn(settings);
            final int maxNumericLength = 12;
            final int maxStringLength = 100; // should suffice

            // this will become second
            debug("adding row column");
            col.setName("ROW");
            col.setNumericLength(maxNumericLength);
            table.addColumnFirst(col);

            col = new TableColumn(settings);

            // this will become first
            debug("adding sheet column");
            col.setName("SHEET");
            col.setStringLength(maxStringLength);
            table.addColumnFirst(col);
        }

        return csvFile != null;
    } // writeCsv

    /**
     * Convert a column number to an Excel column name (A, B, ...).
     *
     * @param columnIndex  starting from 0
     *
     * @return The Excel column name
     */
    private String number2excelColumnName(final int columnIndex) {
        String excelColumnName = "";
        int r = columnIndex + 1; // rest
        int n = 0;
        char c;
        final int nrUpperCaseCharacters = 26;

        while (r > 0) {
            n = (r - 1) % nrUpperCaseCharacters + 1;
            r = (r - n) / nrUpperCaseCharacters;
            c = 'A';
            c += n - 1;
            excelColumnName = Character.toString(c) + excelColumnName;
        }

        return excelColumnName;
    }

    /**
     * Is this sparse array row empty?
     *
     * @param row  The row array
     *
     * @return True or false
     */
    private boolean isEmptyRow(final HashMap<Integer, String> row) {
        for (Integer c : row.keySet()) {
            final String col = row.get(c);

            if (col != null && col.length() > 0) {
                debug("column " + c + " (" + col + ") is not empty");
                return false;
            }
        }
        return true;
    }
}
