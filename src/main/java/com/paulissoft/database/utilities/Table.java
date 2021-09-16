/*
 * Created on Dec 14, 2004
 */
package com.paulissoft.database.utilities;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import java.text.Normalizer;

/**
 * @author Administrator
 * @version 1.0
 */
public class Table {
	
    static String newline="\r\n";

    static String double_quote = "\"";
    
    /**
     * Bad File Name for the Table
     */
    private String badFileName = null;
	
    /**
     * Discard File Name for the Table
     */
    private String discardFileName = null;
	
    /**
     * List of the Table columns
     */
    private List<TableColumn> columns = null;

    /**
     * Comma Separated Values File Extention
     */
    private final String CSV_EXTENTION = ".csv";
	
    /**
     * table bad file extention
     */
    private final String BAD_EXTENTION = ".bad";
	
    /**
     * table discard file extention
     */
    private final String DISCARD_EXTENTION = ".dsc";
	
    /**
     * table log file extention
     */
    private final String LOG_EXTENTION = ".log";
    /**
     * Directory Location referenced by the Table
     */
    private String dirLocation = null;
	
    /**
     * Log File Name for the Table
     */	
    private String logFileName = null;
	
    /**
     * Name of the Table
     */
    private String name = null;
    
    /**
     * Field separator.
     */
    private String fieldSeparator = null;
	
    /**
     * Enclose string.
     */
    private String enclosureString = null;
	
    /**
     * Encoding string.
     */
    private String encoding = null;
	
    private boolean noExternalTable = false;
    
    public Table(String name, String fieldSeparator, String enclosureString, String encoding, boolean noExternalTable) {
        setName(name);
        setFieldSeparator(fieldSeparator);
        setEnclosureString(enclosureString);
        this.encoding = encoding;
        columns = new ArrayList<TableColumn>();
        badFileName = this.name + BAD_EXTENTION;
        discardFileName = this.name + DISCARD_EXTENTION;
        logFileName = this.name + LOG_EXTENTION;
        this.noExternalTable = noExternalTable;
    }

    private String getOracleCharacterSet()
    {
        switch(encoding)
            {
            case "windows-1252":
                return "WE8MSWIN1252";
            case "UTF-8":
                return "AL32UTF8";
            default:
                throw new RuntimeException("Should have encoding of \"windows-1252\" or \"UTF-8\"");
            }
    }
    
    /**
     * 
     * @param column
     */
    public void addColumn(TableColumn column)
    {
        columns.add(column);
    }
    
    /**
     * 
     * @param column
     */
    public void addColumnFirst(TableColumn column)
    {
        columns.add(0, column);
    }
    
    /**
     * 
     * @return column
     */
    public TableColumn getColumn(int i)
    {
        return columns.get(i);
    }

    /**
     * 
     * @return columns
     */
    public int getNrColumns() {
        return columns.size();
    }
	
    /**
     * 
     * @return
     */
    public String getBadFileName() {
        return badFileName;
    }

    /**
     * 
     * @return
     */
    public String getDiscardFileName() {
        return discardFileName;
    }

    /**
     * 
     * @return
     */
	 
    public List getColumns() {
        return columns;
    }
	
    /**
     * Return the SQL String to create the table
     * which references the newly created CSV files derived
     * from the original Excel sheet
     * @return
     */
    public String getDdl()
    {
        String ddl = "CREATE TABLE " +  getName() +newline+"("+ newline;
		
        Iterator iter = columns.iterator();
        while (iter.hasNext()){
            TableColumn c = (TableColumn)iter.next();
            ddl += c.getColumnDdl();
        }
        ddl = ddl.substring(0, ddl.lastIndexOf(","))+ //remove the last comma
            newline+
            ")";

        if (!noExternalTable) {
            ddl += newline
                + "ORGANIZATION EXTERNAL"+ newline
                + "(" + newline
                + "  TYPE oracle_loader" + newline
                + "  DEFAULT DIRECTORY load_dir"+ newline
                + "  ACCESS PARAMETERS "+ newline
                + "  ("+ newline
                + "    RECORDS DELIMITED BY NEWLINE"+ newline
                + "    CHARACTERSET " + getOracleCharacterSet() + newline
                /*
                 * Oracle documentation:
                 *
                 * STRING SIZES ARE IN
                 *
                 * The STRING SIZES ARE IN clause is used to indicate whether the
                 * lengths specified for character strings are in bytes or
                 * characters. If this clause is not specified, then the access
                 * driver uses the mode that the database uses. Character types
                 * with embedded lengths (such as VARCHAR) are also affected by
                 * this clause. If this clause is specified, then the embedded
                 * lengths are a character count, not a byte count. Specifying
                 * STRING SIZES ARE IN CHARACTERS is needed only when loading
                 * multibyte character sets, such as UTF16.
                 */                                                                                                   
                + "    STRING SIZES ARE IN BYTES"+ newline
                + "    FIELD NAMES ALL FILES IGNORE"+ newline
                + "    BADFILE load_dir:'" + getBadFileName() + "'" + newline
                + "    DISCARDFILE load_dir:'" + getDiscardFileName() + "'" + newline
                + "    LOGFILE load_dir:'" + getLogFileName() + "'" + newline
                // Oracle 12C syntax
                + "    FIELDS TERMINATED BY '" + getFieldSeparator() + "' OPTIONALLY ENCLOSED BY '" + getEnclosureString() + "' DATE_FORMAT DATE MASK \"yyyy-mm-dd\""+ newline
                // Oracle 12CR2 syntax
                // + "    FIELDS DATE_FORMAT DATE MASK \"yyyy-mm-dd\" CSV WITHOUT EMBEDDED RECORD TERMINATORS"+ newline
                + "    MISSING FIELD VALUES ARE NULL"+ newline
                + "  )"+ newline
                + "  LOCATION ('"+ getLocation()+"')"+ newline
                +") REJECT LIMIT 0"+newline+newline+newline;
        }

        ddl += ";"+newline;
		
        return ddl;
	
    }

    /**
     * 
     * @return
     */
    public String getDirLocation() {
        return dirLocation;
    }
	
    /**
     * Get the location of the file referenced by the 
     * table.  This will be the same as the 
     * table name with a .csv extention
     * @return
     */
    public String getLocation() {
        return name + CSV_EXTENTION;
    }
	
    /**
     * 
     * @return
     */
    public String getLogFileName() {
        return logFileName;
    }
	
    /**
     * 
     * @return
     */
    public String getName() {
        return (name == null ? null : double_quote + name + double_quote);
    }
	
    /**
     * 
     * @param badFileName
     */
    public void setBadFileName(String badFileName) {
        this.badFileName = badFileName;
    }
	
    /**
     * 
     * @param discardFileName
     */
    public void setDiscardFileName(String discardFileName) {
        this.discardFileName = discardFileName;
    }
	
    /**
     * 
     * @param dirLocation
     */
    public void setDirLocation(String dirLocation) {
        this.dirLocation = dirLocation;
    }
	
    /**
     * 
     * @param logFileName
     */
    public void setLogFileName(String logFileName) {
        this.logFileName = logFileName;
    }
	
    /**
     * 
     * @param name
     */
    public void setName(String name) {
        this.name = Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "").trim();
    }

    public static String getName(String name) {
        return (name == null ? null : double_quote + Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "").trim() + double_quote);
    }

    /**
     * 
     * @param fieldSeparator
     */
    public void setFieldSeparator(String fieldSeparator) {
        this.fieldSeparator = fieldSeparator;
    }

    public String getFieldSeparator() {
        return this.fieldSeparator;
    }

    /**
     * 
     * @param enclosureString
     */
    public void setEnclosureString(String enclosureString) {
        this.enclosureString = enclosureString;
    }

    public String getEnclosureString() {
        return this.enclosureString;
    }
}
