/*
 * Created on Dec 13, 2004  
 */
package com.saternos.database.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
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

/**
 * @author Casimir Saternos
 * @version 1.0
 * 
 *         This program is run at the command line. Given a excel
 *         spreadsheet, a script is generated to create the external table(s)
 *         that references the data from the spreadsheet.
 */
public class ExternalTableGenerator {

    static String newline = "\r\n";
	
    private final String separator = ",";

    private final String enclosure = "\"";
    
    public static void main(String args[]) {
		
        if (args.length != 1) {
            System.out.println(newline + "Usage: ExternalTableGenerator <excel_file_name>" + newline);
            System.exit(0);
        }
        System.out.println("INFO: Begin processing.");
		
        ExternalTableGenerator generator = new ExternalTableGenerator(args[0]);

        System.out.println("INFO: Using working directory " + new File(generator.pwd).getAbsolutePath());
		
        generator.execute();
		
        System.out.println("INFO: Processing complete.");

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
    private List<ExternalTableColumn> externalTables;

    /**
     * Present working directory
     */
    private String pwd;

    /**
     * The Excel Spreadsheet (xls) that is being accessed
     */
    private String spreadsheet;

    private String ddlString ="";
	
    /**
     * @param string
     */
    public ExternalTableGenerator(String spreadsheet) {

        pwd = new File("").getAbsolutePath();
		
        // this.spreadsheet = pwd+File.separator+spreadsheet;
        this.spreadsheet = (new File(spreadsheet)).getAbsolutePath();
		
    }

    /**
     *  Open the specified xls and process it
     */
    private void execute() {

        try {
            ddlString ="CREATE /*OR REPLACE*/ DIRECTORY load_dir AS '"+pwd+"'"+newline+";"+newline+newline; 
			
            Workbook wb;
            
            try {
                POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(spreadsheet));
                wb = new HSSFWorkbook(fs);
            } catch (OfficeXmlFileException e) {
                wb = new XSSFWorkbook(new FileInputStream(spreadsheet));
            }
			
            processWorkbook(wb);
			
            write(ddlString, "ExternalTables.sql");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @param sheet
     * @param table
     */
    private void processSheet(Sheet sheet, ExternalTable table) {
        //Write out a .csv file based upon the sheet
        if (writeCsv(sheet, table)) {
            //Add the ddl for the table to the script
            ddlString += table.getDdl();
        }
    }

    /**
     * @param wb
     * Iterate through each sheet in the workbook
     * and process it
     */
    private void processWorkbook(Workbook wb) {
        
        for (int i = 0; i < wb.getNumberOfSheets(); i++) 
            {
                Sheet sheet = wb.getSheetAt(i);

                System.out.println("INFO: Processing sheet " + i);

                
                ExternalTable table = new ExternalTable(wb.getSheetName(i));
                
                processSheet(sheet, table);
			
                System.out.println("INFO: Table "+ table.getName() + " processed." );
            }
		
    }

    private String getStringValue(Cell cell, ExternalTableColumn col) {
        final String value = cell.getStringCellValue();
        
        // Can not switch from DATE/NUMERIC to STRING

        if (col.getType() == null) {
            col.setType("STRING");
        }
        
        if (col.getType().equals("STRING")) {
            col.setLength(Math.max(col.getLength(), value.length()));
        }

        return value;
    }

    private String getNumericValue(Cell cell, ExternalTableColumn col) {
        String value = null;

        // Can switch STRING to DATE/NUMERIC

        // Test if a date! See https://poi.apache.org/help/faq.html
        if (DateUtil.isCellDateFormatted(cell)) {
            Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

            value = dateFormat.format(date);
            col.setType("DATE");
        } else {
            value = "" + cell.getNumericCellValue();
            col.setNumericPrecision(cell.getNumericCellValue());
            col.setType("NUMERIC");
        }
        
        return value;
    }

    private String getBooleanValue(Cell cell, ExternalTableColumn col) {
        final String value = "" + cell.getBooleanCellValue();
        
        // Treat it as a STRING. Can not switch from DATE/NUMERIC to STRING

        if (col.getType() == null) {
            col.setType("STRING");
        }
        
        if (col.getType().equals("STRING")) {
            col.setLength(Math.max(col.getLength(), value.length()));
        }

        return value;
    }

    /**
     * @param content
     * @param filename
     * Write the given String content to the file system
     * using the String filename specified
     */
    private void write(String content, String filename) {

        try {
            File f = new File(filename);
            f.createNewFile();
            FileWriter fr = new FileWriter(filename);
            fr.write(content);
            fr.flush();
            fr.close();
			
            System.out.println("INFO: File " + filename + " created.");
			
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @param sheet
     * @param table
     */
    private Boolean writeCsv(Sheet sheet, ExternalTable table) {

        Row names = sheet.getRow(COLUMN_NAME_ROW);
        Row types = sheet.getRow(COLUMN_TYPE_ROW);

        ArrayList<ExternalTableColumn> cols = new ArrayList<ExternalTableColumn>();

        String csv = "";

        Iterator<Row> rowIterator = sheet.rowIterator();
            
        //skip putting the column names and type length row in the csv
        for (int r = COLUMN_NAME_ROW; rowIterator.hasNext(); r++) {

            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            
            for (short c = 0; cellIterator.hasNext(); c++) {
                try {
                
                    Cell cell = cellIterator.next();

                    ExternalTableColumn col = (r == COLUMN_NAME_ROW ? new ExternalTableColumn() : cols.get(c));
                    
                    String value = null;

                    if (cell == null) continue;

                    switch(r)
                        {
                        case COLUMN_NAME_ROW:
                            value = cell.getStringCellValue();
                            col.setName(value);
                            cols.add(col);
                            break;
                    
                        default:
                            // column type can switch from string to numeric but not vice versa
                            switch(cell.getCellType())
                                {
                                case FORMULA:
                                    try {
                                        value = getStringValue(cell, col);
                                        // string?
                                    } catch (IllegalStateException e1) {
                                        // java.lang.IllegalStateException: Cannot get a STRING value from a NUMERIC cell

                                        try {
                                            value = getNumericValue(cell, col);
                                        } catch (IllegalStateException e2) {
                                            // java.lang.IllegalStateException: Cannot get a NUMERIC value from a BOOLEAN formula cell
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
		
        // Set the table definition information
        table.setColumns(cols);

        if (csv.length() > 0) {
            // Final newline causes problems so remove it
            write(csv.substring(0, csv.length()-1), table.getLocation());

            return true;
        } else {
            System.out.println("WARNING: Sheet does not contain data");
            
            return false;
        }
    }
}
