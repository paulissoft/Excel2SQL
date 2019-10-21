/*
 * Created on Dec 13, 2004  
 */
package com.saternos.database.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

import org.apache.poi.ss.usermodel.Workbook; // interface
import org.apache.poi.ss.usermodel.Sheet; // interface
import org.apache.poi.ss.usermodel.Row; // interface
import org.apache.poi.ss.usermodel.Cell; // interface
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

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

    @Parameter(names = "-verbose", description = "Level of verbosity")
    private Integer verbose = 1;

    @Parameter(names = "-sheet", description = "Sheet name regular expression")
    private String sheetName = ".*";

    /**
     * The Excel Spreadsheets (xls) that are being accessed
     */
    @Parameter(description = "Spreadsheets")
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
  
    public static void main(String ... args) {
    
        if (args.length == 0) {
            System.out.println(newline + "Usage: ExternalTableGenerator <excel_file_name 1> .. <excel_file_name N>" + newline);
            System.exit(0);
        }
        
        ExternalTableGenerator generator = new ExternalTableGenerator();

        JCommander.newBuilder()
            .addObject(generator)
            .build()
            .parse(args);

        generator.execute();
    }
  
    // All sheets in the workbook use the following constants.
  
    /**
     * The index of the row where the values that are used
     * for the names of the table columns are retrieved
     */
    private final int COLUMN_NAME_ROW = 0;

    /**
     * The index of the row where the values that are used
     * for the types of the table columns are retrieved
     * (VARCHAR2 and NUMBER are the only two types currently
     * handled).
     */
    private final int COLUMN_TYPE_ROW = 1;

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
     *  Open the specified xls and process it
     */
    private void execute() {
        info("Begin processing.");

        info("Using working directory " + new File(pwd).getAbsolutePath());

        ddlString = "CREATE /*OR REPLACE*/ DIRECTORY load_dir AS '"+pwd+"'"+newline+";"+newline+newline; 

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

                info("Processing workbook " + spreadsheet);

                processWorkbook(wb, i == 0, i == spreadsheets.size() - 1);                
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        write(ddlString, "ExternalTables.sql", false, false);
            
        info("Processing complete.");
    }

    /**
     * @param sheet
     * @param table
     */
    private void processSheet(Sheet sheet, ExternalTable table, boolean first, boolean last) {
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
    private void processWorkbook(Workbook wb, boolean first, boolean last) {

        //if (!first && !(externalTables.size() == wb.getNumberOfSheets())) {
        //    throw new RuntimeException("External tables size (" + externalTables.size() + ") should be equal to the number of sheets (" + wb.getNumberOfSheets() + ")");
        //}

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);

            if (!wb.getSheetName(i).matches(sheetName)) {
                info("Skipping sheet " + i + ": " + wb.getSheetName(i) + " since it does not match " + sheetName);
                continue;
            } else {
                info("Processing sheet " + i + ": " + wb.getSheetName(i));
            }
                
            ExternalTable table = null;

            if (first) {
                table = new ExternalTable(wb.getSheetName(i));
                externalTables.add(table);
            } else {
                final String tableName = ExternalTable.getName(wb.getSheetName(i));

                for (int index = 0; index < externalTables.size(); index++) {
                    if (externalTables.get(index).getName().equals(tableName)) {
                        table = externalTables.get(index);
                        break;
                    }
                }
                
                if (table == null) {
                    throw new RuntimeException("Could not find table name (" + tableName + ")");
                }
            }

            
            processSheet(sheet, table, first, last);

            if (last) {      
                info("Table "+ table.getName() + " processed." );
            }
        }        
    }

    private String getStringValue(Cell cell, ExternalTableColumn col) {
        final String value = cell.getStringCellValue();
        
        col.setStringLength(value.length());

        if (value != null && value.length() > 0) {
            debug("getStringValue for column '" + col.getName() + "' and value '" + value + "'" );
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
        } else {
            value = "" + cell.getNumericCellValue();
            
            // store the length first since it may be important in setType()
            col.setNumericLength(value.length());
            col.setNumericPrecision(cell.getNumericCellValue());
        }

        debug("getNumericValue for column '" + col.getName() + "' and value '" + value + "'" );

        return value;
    }

    private String getBooleanValue(Cell cell, ExternalTableColumn col) {
        final String value = "" + cell.getBooleanCellValue();
        
        // Treat it as a STRING.
        col.setStringLength(value.length());

        debug("getBooleanValue for column '" + col.getName() + "' and value '" + value + "'" );

        return value;
    }

    /**
     * @param content
     * @param filename
     * Write the given String content to the file system
     * using the String filename specified
     */
    private void write(String content, String filename, Boolean utf8, boolean append) {

        try {
            // GJP 2019-10-18  Seems not necessary.
            // File f = new File(filename);
            // f.createNewFile();

            Writer fr;

            if ( !utf8 ) {
                fr = new OutputStreamWriter(new FileOutputStream(filename, append), "windows-1252");
            } else {
                fr = new OutputStreamWriter(new FileOutputStream(filename, append), StandardCharsets.UTF_8);
            }
            
            fr.write(content);
            fr.flush();
            fr.close();
            
            info("Finished " + (append ? "appending to" : "creating") + " " + filename);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @param sheet
     * @param table
     * @param first  First workbook
     * @param last   Last workbook
     */
    private Boolean writeCsv(Sheet sheet, ExternalTable table, boolean first, boolean last) {

        Row names = sheet.getRow(COLUMN_NAME_ROW);
        Row types = sheet.getRow(COLUMN_TYPE_ROW);

        String csv = "";
        String csvEmptyRow = "";
        String progress = null;

        Iterator<Row> rowIterator = sheet.rowIterator();
            
        for (int r = COLUMN_NAME_ROW; rowIterator.hasNext(); r++) {

            debug("Processing row " + (r+1));
            
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
            
            for (short c = 0; (r == COLUMN_NAME_ROW || c < table.getNrColumns()); c++) {
                debug("Processing column " + (c+1));
                            
                try {                
                    Cell cell = cellIterator.hasNext() ? cellIterator.next() : null;
                    
                    if (cell == null) {
                        debug("No cell defined");

                        if (r == COLUMN_NAME_ROW) {
                            break;
                        } else {
                            csvRow += ExternalTable.SEPARATOR;
                            continue;
                        }
                    }

                    assert(cell != null);
                    
                    debug("cell address: " + cell.getAddress() + "; cell column index: " + cell.getColumnIndex());

                    // Sometimes there may be cells missing so after cell column index 0 may come cell column index 2
                    // But not for the header
                    if (r == COLUMN_NAME_ROW && !(c == cell.getColumnIndex())) {
                        throw new RuntimeException("There should be no columns missing for the header");
                    }
                                        
                    String value = null;
                    String missingColumns = "";
                    
                    for ( ; c < cell.getColumnIndex(); c++ ) {
                        missingColumns += ExternalTable.SEPARATOR;
                    }
                    
                    assert(c == cell.getColumnIndex());
                    
                    ExternalTableColumn col = (r == COLUMN_NAME_ROW && first ? new ExternalTableColumn() : table.getColumn(c));                    

                    switch(r)
                        {
                        case COLUMN_NAME_ROW:
                            // Some names are just numbers, strangely enough (column name 14)
                            try {
                                value = cell.getStringCellValue();
                                // string?
                            } catch (IllegalStateException e1) {
                                // java.lang.IllegalStateException: Cannot get a STRING value from a NUMERIC cell
                                value = "" + cell.getNumericCellValue();
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
                            csvEmptyRow += ExternalTable.SEPARATOR;                            
                            break;
                    
                        default:
                            // column type can switch from string to numeric but not vice versa
                            switch(cell.getCellType())
                                {
                                case FORMULA:
                                    // try to be the most specific: booleans, numbers and then strings
                                    try {
                                        value = getBooleanValue(cell, col);
                                        // string?
                                    } catch (IllegalStateException e1) {
                                        // java.lang.IllegalStateException: Cannot get a BOOLEAN value from a NUMERIC formula cell

                                        try {
                                            value = getNumericValue(cell, col);
                                        } catch (IllegalStateException e2) {
                                            // java.lang.IllegalStateException: Cannot get a NUMERIC value from a STRING cell
                                            value = getStringValue(cell, col);                                            
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

                    // The Column Name row is only prin,ted the first time else it is used for verification
                    if (r == COLUMN_NAME_ROW && !first) continue;                        

                    // see https://en.wikipedia.org/wiki/Comma-separated_values
                    value.replace(ExternalTable.ENCLOSURE, ExternalTable.ENCLOSURE + ExternalTable.ENCLOSURE);
                    if (value.contains(ExternalTable.ENCLOSURE) || value.contains(ExternalTable.SEPARATOR)) {
                        value = ExternalTable.ENCLOSURE + value + ExternalTable.ENCLOSURE;
                    }
                    csvRow += missingColumns + value + ExternalTable.SEPARATOR;
                } catch (Exception e) {
                    System.err.println("Error in line " + (r+1) + " for column " + (c+1));
                    throw e;
                }
            }
            if (!csvRow.equals(csvEmptyRow)) {
                csv += csvRow + newline;
            } else {
                info("Skipping row " + (r+1) + " since it is empty");
            }
        }
        
        if (progress != null) {
            info(progress);
        }
                    
        System.out.println("");

        if (csv.length() > 0) {
            // Final newline causes problems so remove it
            write(csv.substring(0, csv.length()-1), table.getLocation(), false, !first);

            return true;
        } else {
            System.out.println("WARNING: Sheet does not contain data");
            
            return false;
        }
    }
}
