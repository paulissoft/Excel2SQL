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

    private Settings settings = null;
    
    public Table(String name, Settings settings) {
        setName(name);
        this.settings = settings;
        columns = new ArrayList<TableColumn>();
        badFileName = this.name + BAD_EXTENTION;
        discardFileName = this.name + DISCARD_EXTENTION;
        logFileName = this.name + LOG_EXTENTION;
    }

    private String getOracleCharacterSet()
    {
        switch(getEncoding())
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
        String ddl = "CREATE TABLE " + (settings.sqlDatabase.equals(Settings.POSTGRESQL) ? "IF NOT EXISTS " : "") + getName() + Settings.NL + "(" + Settings.NL;
		
        Iterator iter = columns.iterator();
        while (iter.hasNext()){
            TableColumn c = (TableColumn)iter.next();
            ddl += c.getColumnDdl();
        }
        ddl = ddl.substring(0, ddl.lastIndexOf(","))+ //remove the last comma
            Settings.NL+
            ")";

        if (settings.sqlDatabase.equals(Settings.ORACLE)) {
            ddl += Settings.NL
                + "ORGANIZATION EXTERNAL"+ Settings.NL
                + "(" + Settings.NL
                + "  TYPE oracle_loader" + Settings.NL
                + "  DEFAULT DIRECTORY load_dir"+ Settings.NL
                + "  ACCESS PARAMETERS "+ Settings.NL
                + "  ("+ Settings.NL
                + "    RECORDS DELIMITED BY SETTINGS.NL"+ Settings.NL
                + "    CHARACTERSET " + getOracleCharacterSet() + Settings.NL
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
                + "    STRING SIZES ARE IN BYTES"+ Settings.NL
                + "    FIELD NAMES ALL FILES IGNORE"+ Settings.NL
                + "    BADFILE load_dir:'" + getBadFileName() + "'" + Settings.NL
                + "    DISCARDFILE load_dir:'" + getDiscardFileName() + "'" + Settings.NL
                + "    LOGFILE load_dir:'" + getLogFileName() + "'" + Settings.NL
                // Oracle 12C syntax
                + "    FIELDS TERMINATED BY '" + getFieldSeparator() + "' OPTIONALLY ENCLOSED BY '" + getEnclosureString() + "' DATE_FORMAT DATE MASK \"yyyy-mm-dd\""+ Settings.NL
                // Oracle 12CR2 syntax
                // + "    FIELDS DATE_FORMAT DATE MASK \"yyyy-mm-dd\" CSV WITHOUT EMBEDDED RECORD TERMINATORS"+ Settings.NL
                + "    MISSING FIELD VALUES ARE NULL"+ Settings.NL
                + "  )"+ Settings.NL
                + "  LOCATION ('"+ getLocation()+"')"+ Settings.NL
                +") REJECT LIMIT 0"+Settings.NL+Settings.NL+Settings.NL;
        }

        ddl += ";"+Settings.NL;
		
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
        return (name == null ? null : Settings.QQ + name + Settings.QQ);
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
        return (name == null ? null : Settings.QQ + Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "").trim() + Settings.QQ);
    }

    /**
     * 
     * @param fieldSeparator
     */
    public String getFieldSeparator() {
        return this.settings.columnSeparator;
    }

    /**
     * 
     * @param enclosureString
     */
    public String getEnclosureString() {
        return this.settings.enclosureString;
    }

    public String getEncoding() {
        return this.settings.encoding;
    }

    static protected String preamble(Settings settings) {
        if (settings.sqlDatabase.equals(Settings.ORACLE)) {
            return "CREATE /*OR REPLACE*/ DIRECTORY load_dir AS '" + Settings.PWD + "'" + Settings.NL + ";" + Settings.NL + Settings.NL;
        } else {
            return "";
        }
    }
}
