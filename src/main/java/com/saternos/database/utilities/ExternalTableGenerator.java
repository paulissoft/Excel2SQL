/*
 * Created on Dec 13, 2004  
 */
package com.saternos.database.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.io.FileWriter;
import java.nio.charset.StandardCharsets;
import java.io.Writer;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.CommonObjectDataSubRecord;
import org.apache.poi.hssf.record.ObjRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SubRecord;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook; // interface
import org.apache.poi.ss.usermodel.Sheet; // interface
import org.apache.poi.ss.usermodel.Row; // interface
import org.apache.poi.ss.usermodel.Cell; // interface
import org.apache.poi.ss.usermodel.CellStyle; // interface
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.beust.jcommander.Parameter;
import com.beust.jcommander.JCommander;

/**
 * @author Casimir Saternos
 * @version 1.0
 * 
 *         This program is run at the command line. Given one or more excel
 *         spreadsheets, a script is generated to create the external table(s)
 *         that references the data from the spreadsheets. 
 *
 *         When multiple spreadsheets are supplied, 
 *         they must all have the same number and names of sheets.
 */
public class ExternalTableGenerator {

    @Parameter(names = "--verbose", description = "Level of verbosity")
    private Integer verbose = 1;

    @Parameter(names = "--sheet-name-expression", description = "Sheet name(s) must match this regular expression")
    private String sheetNameExpression = ".*";

    @Parameter(names = "--sql-table-name", description = "A list of SQL table name(s) to use instead of the sheet name(s)")
    private List<String> tableNames = new ArrayList<>();

    @Parameter(names = "--column-separator", description = "The column separator")
    private String columnSeparator = ",";

    @Parameter(names = "--enclosure-string", description = "The enclosure string")
    private String enclosureString = "\"";

    @Parameter(names = "--encoding", description = "The encoding to use (default \"windows-1252\")")
    private String encoding = "windows-1252";
    
    @Parameter(names = "--write-bom", description = "Write the BOM at the beginning of the file.")
    private boolean writeBOM = false;

    @Parameter(names = "--no-header", description = "The first row does NOT contain the column names")
    private boolean noHeader = false;

    @Parameter(names = "--no-external-table", description = "Do not create a script for an external table just for a table")
    private boolean noExternalTable = false;

    @Parameter(names = "--add-row-number", description = "Add the sheet row number (starting from 1) in the CSV?")
    private boolean addRowNumber = false;

    @Parameter(names = { "--help", "-h" }, description = "This help", help = true)
    private boolean help;
    
    /**
     * The Excel Spreadsheets (.xls or .xlsx) that are being accessed
     */
    @Parameter(description = "spreadsheet...", required = true)
    private List<String> spreadsheets = new ArrayList<String>();

    private void info(String str) {
        if (verbose >= 1) {
            System.out.println("INFO: " + str);
        }
    }

    private void debug(String str) {
        if (verbose >= 2) {
            System.out.println("DEBUG: " + str);
        }
    }
    
    static String newline = "\r\n";

    private DataFormatter dataFormatter = new DataFormatter();

    private FormulaEvaluator formulaEvaluator = null;
  
    public static void main(String ... args) throws java.io.IOException {
        ExternalTableGenerator generator = new ExternalTableGenerator();

        JCommander jc = JCommander.newBuilder()
            .addObject(generator)
            .build();

        jc.setProgramName("ExternalTableGenerator");

        try {
            jc.parse(args);

            if (generator.help) {
                jc.usage();
            } else {
                // Check file exists as a regular file
                for (int i = 0; i < generator.spreadsheets.size(); i++) {
                    File f = new File(generator.spreadsheets.get(i));

                    try {
                        assert(f.exists() && !f.isDirectory());
                    } catch(AssertionError e) {
                        System.err.println("File '" + generator.spreadsheets.get(i) + "' does not exist or is not a regular file");
                        throw e;
                    }
                }
            }
        } catch(Exception e) {
            jc.usage();
            throw e;
        }

        if (!generator.help) {
            generator.execute();
        }
    }
  
    // All sheets in the workbook use the following constants.
  
    /**
     * The index of the row where the values that are used
     * for the names of the table columns are retrieved.
     * But only if !noHeader
     */
    private final int COLUMN_NAME_ROW = 0;

    /**
     * List of external table definitions
     */
    // private List<ExternalTableColumn> externalTables;
    private List<ExternalTable> externalTables;


    /**
     * Present working directory
     */
    private String pwd;

    private String ddlString = "";
  
    public ExternalTableGenerator() {
        this.pwd = new File("").getAbsolutePath();    
        this.externalTables = new ArrayList<ExternalTable>();
    }

    /**
     *  Open the specified .xls or .xlsx and process it
     */
    private void execute() throws java.io.IOException {
        info("Begin processing.");

        info("Using working directory " + new File(pwd).getAbsolutePath());

        if (!this.noExternalTable) {
            ddlString = "CREATE /*OR REPLACE*/ DIRECTORY load_dir AS '"+pwd+"'"+newline+";"+newline+newline;
        } else {
            ddlString = "";
        }

        for (int i = 0; i < spreadsheets.size(); i++) {
            final String spreadsheet = (new File(spreadsheets.get(i))).getAbsolutePath();

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
                            switch(c.getCellType()) {
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

                processWorkbook(wb, i == 0, i == spreadsheets.size() - 1);

                debugWorkbook(new FileInputStream(spreadsheet));
            } catch (Exception e) {
                e.printStackTrace();
                
                throw e;
            }
        }

        write(ddlString, (!this.noExternalTable ? "External" : "") + "Tables.sql", "UTF-8", false, false);
            
        info("Processing complete.");
    }

    /**
     * @param sheet
     * @param table
     */
    private void processSheet(Sheet sheet, ExternalTable table, boolean first, boolean last) throws java.io.IOException {
        //Write out a .csv file based upon the sheet
        if (writeCsv(sheet, table, first, last) && last) {
            // Add the ddl for the table to the script
            ddlString += table.getDdl();
        }
    }

    /**
     * @param wb
     * Iterate through each sheet in the workbook
     * and process it
     */
    private void processWorkbook(Workbook wb, boolean first, boolean last) throws java.io.IOException {

        //if (!first && !(externalTables.size() == wb.getNumberOfSheets())) {
        //    throw new RuntimeException("External tables size (" + externalTables.size() + ") should be equal to the number of sheets (" + wb.getNumberOfSheets() + ")");
        //}

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);

            if (!wb.getSheetName(i).matches(sheetNameExpression)) {
                info("Skipping sheet " + i + ": " + wb.getSheetName(i) + " since it does not match " + sheetNameExpression);
                continue;
            } else {
                info("Processing sheet " + i + ": " + wb.getSheetName(i));
            }
                
            ExternalTable table = null;
            final String tableName = ( i < tableNames.size() ? tableNames.get(i) : wb.getSheetName(i) );

            if (first) {
                table = new ExternalTable(tableName, columnSeparator, enclosureString, encoding, noExternalTable);
                externalTables.add(table);
            } else {
                final String sqlTableName = ExternalTable.getName(tableName);

                for (int index = 0; index < externalTables.size(); index++) {
                    if (externalTables.get(index).getName().equals(sqlTableName)) {
                        table = externalTables.get(index);
                        break;
                    }
                }
                
                if (table == null) {
                    throw new RuntimeException("Could not find table name (" + sqlTableName + ")");
                }
            }

            
            processSheet(sheet, table, first, last);

            if (last) {      
                info("Table "+ table.getName() + " processed." );
            }
        }        
    }

    private void debugWorkbook(InputStream spreadsheet) throws java.io.IOException {
        // GJP 2021-09-15 Check Boolean cells
                
        HSSFRequest req = new HSSFRequest();
        req.addListenerForAllRecords(new ProcessFile());
        HSSFEventFactory factory = new HSSFEventFactory();
        factory.processEvents(req, spreadsheet);
    }

    private String getStringValue(Cell cell, ExternalTableColumn col) {
        // final String value = cell.getStringCellValue();
        final String value = cell.getRichStringCellValue().getString();        
        
        col.setStringLength(value.length());

        if (value != null && value.length() > 0) {
            debug( "getStringValue for column '" + col.getName() +
                   "', value '" + value + "'" +
                   "', length " + col.getStringLength() +
                   " and format " + cell.getCellStyle().getDataFormatString() );
        }
        
        return value;
    }

    private String getNumericValue(Cell cell, ExternalTableColumn col) {
        String value = null;

        // Test if a date! See https://poi.apache.org/help/faq.html
        final Boolean isDate = DateUtil.isCellDateFormatted(cell);
        
        if (isDate) {
            Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

            value = dateFormat.format(date);
            
            col.setDateLength(value.length());
            
            debug( "getNumericValue for date column '" + col.getName() +
                   "', value '" + value +
                   "', length " + col.getDateLength() +
                   " and format " + cell.getCellStyle().getDataFormatString() );
        } else {
            // Remove trailing % and remove thousands separator (comma)
            value = dataFormatter.formatCellValue(cell, formulaEvaluator).replace(",", "");

            if (value.endsWith("%")) {
                // replace the percent symbol and divide by 100
                value = (new Double(Double.valueOf(value.replace("%", "")) / 100)).toString();
            }

            String[] parts = value.split("\\.");

            for (int i = 0; i < parts.length; i++) {
                debug("parts[" + i + "]= '" + parts[i] + "'");
            }

            // store the length first since it may be important in setType()
            col.setNumericLength((parts.length == 1 ? parts[0].length() : parts[0].length() + parts[1].length()));
            col.setNumericPrecision((parts.length == 1 ? 0 : parts[1].length()));

            debug( "getNumericValue for numeric column '" + col.getName() +
                   "', value '" + value +
                   "', length " + col.getNumericLength() +
                   ", precision " + col.getNumericPrecision() +
                   " and format " + cell.getCellStyle().getDataFormatString() );
        }

        return value;
    }

    private String getBooleanValue(Cell cell, ExternalTableColumn col) {
        final String value = "" + cell.getBooleanCellValue();
        
        // Treat it as a STRING.
        col.setStringLength(value.length());

        debug("getBooleanValue for column '" + col.getName() + "' and value '" + value + "'" );
        debug("cell: " + cell.toString());

        return value;
    }

    /**
     * @param content
     * @param filename
     * @param encoding
     * @param writeBOM
     * @param append
     *
     * Write the given String content to the file system
     * using the String filename specified
     */
    private void write(String content, String filename, String encoding, boolean writeBOM, boolean append) throws java.io.IOException {

        try {
            // GJP 2019-10-18  Seems not necessary.
            // File f = new File(filename);
            // f.createNewFile();

            PrintStream out = new PrintStream(new FileOutputStream(filename, append), false, encoding);

            // write the BOM?
            if (!append && encoding.equals("UTF-8") && writeBOM) {
                byte[] bom = {(byte)0xEF, (byte)0xBB, (byte)0xBF};
                out.write(bom);
            }

            out.print(content);
            out.close();
            
            info("Finished " + (append ? "appending to" : "creating") + " " + filename);
        } catch (Exception e) {
            e.printStackTrace();
            
            throw e;
        }
    }

    /**
     * @param sheet
     * @param table
     * @param first  First workbook
     * @param last   Last workbook
     */
    private Boolean writeCsv(Sheet sheet, ExternalTable table, boolean first, boolean last) throws java.io.IOException {

        // Row names = sheet.getRow(COLUMN_NAME_ROW);

        String csv = "";
        String progress = null;

        Iterator<Row> rowIterator = sheet.rowIterator();
            
        for (int r = COLUMN_NAME_ROW; rowIterator.hasNext(); r++) {

            debug("Processing Excel row " + (r+1));
            
            switch (r % 10)
                {
                case 0:
                    progress = "Processing row " + (r+1);
                    break;
                    
                case 9:
                    progress += ".";
                    info(progress);
                    progress = null;
                    break;
                    
                default:
                    progress += ".";
                }

            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            boolean rowEmpty = true;
            String csvRow = "";

            debug("row number          : " + row.getRowNum());
            debug("first cell number   : " + row.getFirstCellNum());
            debug("last cell number + 1: " + row.getLastCellNum());

            // process column if
            // 1) there is no header (we may always add column names) OR
            // 2) this is the column name row OR
            // 3) the column is part of the columns found
            for (short c = 0; (this.noHeader || r == COLUMN_NAME_ROW || c < table.getNrColumns()); c++) {
                debug("Processing Excel column " + (c+1));
                debug("Number of table columns: " + table.getNrColumns());
                            
                try {                
                    Cell cell = cellIterator.hasNext() ? cellIterator.next() : null;
                    
                    if (cell == null) {
                        debug("No cell defined");

                        if (this.noHeader || r == COLUMN_NAME_ROW) {
                            break; // no header column to add
                        } else {
                            csvRow += table.getFieldSeparator();
                            continue;
                        }
                    }

                    assert(cell != null);
                    
                    debug("cell address: " + cell.getAddress() + "; cell column index: " + cell.getColumnIndex());

                    // Sometimes there may be cells missing so after cell column index 0 may come cell column index 2.
                    // But not for the header!
                    if (!this.noHeader && r == COLUMN_NAME_ROW && !(c == cell.getColumnIndex())) {
                        throw new RuntimeException("There should be no columns missing for the header");
                    }
                                        
                    String value = null;
                    String missingColumns = "";

                    if (this.noHeader) {
                        // a data row when there is no header: add missing columns
                        for ( ; c < cell.getColumnIndex(); c++ ) {
                            // add this column as a header column?
                            if (c >= table.getNrColumns()) {
                                ExternalTableColumn col = new ExternalTableColumn();

                                debug("adding column " + (c+1) + " as header (1)");
                                col.setName(number2excelColumnName(c+1));
                                table.addColumn(col);
                            }
                            missingColumns += table.getFieldSeparator();
                        }
                        assert(c == cell.getColumnIndex());
                        
                        // add this column as a header column?
                        if (c >= table.getNrColumns()) {
                            ExternalTableColumn col = new ExternalTableColumn();

                            debug("adding column " + (c+1) + " as header (2)");
                            col.setName(number2excelColumnName(c+1));
                            table.addColumn(col);
                        }
                        
                    } else if (r != COLUMN_NAME_ROW) {
                        // a data row when there is a header: add missing columns
                        for ( ; c < Math.min(table.getNrColumns()-1, cell.getColumnIndex()); c++ ) {
                            missingColumns += table.getFieldSeparator();
                        }
                        assert(c == cell.getColumnIndex() || c == table.getNrColumns()-1);
                    } else {
                        assert(c == cell.getColumnIndex());
                    }

                    if (c != cell.getColumnIndex()) {
                        // the cell is beyond the number of heading cells
                        value = "";
                    } else {
                        debug("Processing Excel column " + (c+1));
                        
                        // new column if 
                        ExternalTableColumn col = (!this.noHeader && r == COLUMN_NAME_ROW && first ? new ExternalTableColumn() : table.getColumn(c));                    

                        if (!this.noHeader && r == COLUMN_NAME_ROW) {
                            // Some names are just numbers, strangely enough (column name 14)
                            try {
                                // value = cell.getStringCellValue();
                                value = cell.getRichStringCellValue().getString();
                                // string?
                            } catch (IllegalStateException e1) {
                                // java.lang.IllegalStateException: Cannot get a STRING value from a NUMERIC cell
                                value = dataFormatter.formatCellValue(cell);
                            }
                            
                            info((first ? "Scanning" : "Skipping") + " heading " + (c+1) + ": " + value);

                            if (first) {
                                col.setName(value);
                                table.addColumn(col);
                            } else {
                                if (!col.getName().equals(ExternalTable.getName(value))) {
                                    throw new RuntimeException("Column name (" + col.getName() + ") should be equal to the external table name (" + ExternalTable.getName(value) + ")"); // check column name
                                }
                            }
                        } else {
                            // column type can switch from string to numeric but not vice versa
                            switch(cell.getCellType())
                                {
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
                        }
                    }

                    // The Column Name row is only printed the first time else it is used for verification
                    if (!this.noHeader && r == COLUMN_NAME_ROW && !first) continue;

                    // see https://en.wikipedia.org/wiki/Comma-separated_values
                    value.replace(table.getEnclosureString(), table.getEnclosureString() + table.getEnclosureString());
                    if (value.contains(table.getEnclosureString()) || value.contains(table.getFieldSeparator())) {
                        value = table.getEnclosureString() + value + table.getEnclosureString();
                    }
                    csvRow += missingColumns + value + table.getFieldSeparator();
                } catch (Exception e) {
                    System.err.println("Error in line " + (r+1) + " for column " + (c+1));
                    throw e;
                }
            }
            if (!isEmptyRow(csvRow)) {
                csv += (this.addRowNumber ? "" + (row.getRowNum()+1) + table.getFieldSeparator() : "" ) + csvRow + newline;
            } else {
                info("Skipping row " + (r+1) + " since it is empty");
            }
        }
        
        if (progress != null) {
            info(progress);
        }
                    
        System.out.println("");

        if (this.addRowNumber) {
            ExternalTableColumn col = new ExternalTableColumn();

            debug("adding row column");
            col.setName("ROW");
            col.setNumericLength(12);
            table.addColumnFirst(col);            
        }

        if (csv.length() > 0) {
            // Final newline causes problems so remove it
            write(csv.substring(0, csv.length()-1), table.getLocation(), this.encoding, this.writeBOM, !first);

            return true;
        } else {
            System.out.println("WARNING: Sheet does not contain data");
            
            return false;
        }
    }

    // columnIndex starting from 1
    private String number2excelColumnName(int columnIndex) {
        assert(columnIndex > 0);
        
        String excelColumnName = "";
        int r = columnIndex; // rest
        int n = 0;
        final char A = 'A';
        char c;
        
        while (r > 0) {
            n = (r - 1) % 26 + 1;
            r = (r - n) / 26;
            c = A;
            c += n - 1;
            excelColumnName = Character.toString(c) + excelColumnName;
        }

        return excelColumnName;
    }

    private boolean isEmptyRow(String row) {
        return row.replace(this.columnSeparator, "").length() == 0;
    }
}


class ProcessFile implements HSSFListener {

    public void processRecord(Record record) {
        switch (record.getSid()) {
        case ObjRecord.sid:
            ObjRecord objRec = (ObjRecord) record;
            List<SubRecord> subRecords = objRec.getSubRecords();
            for (SubRecord subRecord : subRecords) {
                if (subRecord instanceof CommonObjectDataSubRecord) {
                    CommonObjectDataSubRecord datasubRecord = (CommonObjectDataSubRecord) subRecord;
                    if (datasubRecord.getObjectType() == CommonObjectDataSubRecord.OBJECT_TYPE_CHECKBOX) {
                        System.out.println("ObjId: "
                                           + datasubRecord.getObjectId() +
                                           "Details: "
                                           + datasubRecord.toString());
                    }
                }
            }
            break;
        }
    }
}
