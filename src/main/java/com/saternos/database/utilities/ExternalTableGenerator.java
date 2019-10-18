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
  
    private final String separator = ",";

    private final String enclosure = "\"";
    
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
    private List<List<ExternalTableColumn>> externalTableColumns;


    /**
     * Present working directory
     */
    private String pwd;

    private String ddlString = "";
  
    public ExternalTableGenerator() {
        this.pwd = new File("").getAbsolutePath();    
        this.externalTables = new ArrayList<ExternalTable>();
        this.externalTableColumns = new ArrayList<List<ExternalTableColumn>>();
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
    private void processSheet(Sheet sheet, ExternalTable table, List<ExternalTableColumn> cols, boolean first, boolean last) {
        //Write out a .csv file based upon the sheet
        if (writeCsv(sheet, table, cols, first, last) && last) {
            // Set the table definition information
            table.setColumns(cols);
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

        if (!first) {
            assert(externalTables.size() == wb.getNumberOfSheets());
            assert(externalTableColumns.size() == wb.getNumberOfSheets());
        }

        for (int i = 0; i < wb.getNumberOfSheets(); i++) 
            {
                Sheet sheet = wb.getSheetAt(i);

                info("Processing sheet " + i + ": " + wb.getSheetName(i));
                
                ExternalTable table;
                List<ExternalTableColumn> cols;

                if (first) {
                    table = new ExternalTable(wb.getSheetName(i));
                    externalTables.add(i, table);
                    cols = new ArrayList<ExternalTableColumn>();
                    externalTableColumns.add(i, cols);
                } else {
                    table = externalTables.get(i);
                    cols = externalTableColumns.get(i);
                }

                processSheet(sheet, table, cols, first, last);

                if (last) {      
                    info("Table "+ table.getName() + " processed." );
                }
            }
    
    }

    private String getStringValue(Cell cell, ExternalTableColumn col) {
        final String value = cell.getStringCellValue();
        
        col.setStringLength(value.length());

        debug("getStringValue for column '" + col.getName() + "' and value '" + value + "'" );
        
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
    private Boolean writeCsv(Sheet sheet, ExternalTable table, List<ExternalTableColumn> cols, boolean first, boolean last) {

        Row names = sheet.getRow(COLUMN_NAME_ROW);
        Row types = sheet.getRow(COLUMN_TYPE_ROW);

        String csv = "";
        String progress = null;

        Iterator<Row> rowIterator = sheet.rowIterator();
            
        for (int r = COLUMN_NAME_ROW; rowIterator.hasNext(); r++) {

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
            
            for (short c = 0; cellIterator.hasNext() && (r == COLUMN_NAME_ROW || c < cols.size()); c++) {
                try {
                
                    Cell cell = cellIterator.next();
                    
                    if (cell == null) continue;

                    ExternalTableColumn col = (r == COLUMN_NAME_ROW && first ? new ExternalTableColumn() : cols.get(c));
                    
                    String value = null;

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
                            
                            info("Scanning heading " + (c+1) + ": " + value);

                            if (first) {
                                col.setName(value);
                                cols.add(col);
                            } else {
                                info("Scanning column " + (c+1) + ": " + col.getName());
                                assert(col.getName().equals(ExternalTable.getName(value))); // check column name
                            }
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
                    value.replace(enclosure, enclosure + enclosure);
                    if (value.contains(enclosure) || value.contains(separator)) {
                        value = enclosure + value + enclosure;
                    }
                    csv += value + separator;
                } catch (Exception e) {
                    System.err.println("Error in line " + (r+1) + " for column " + (c+1));
                    throw e;
                }
            }
            csv += newline;
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
