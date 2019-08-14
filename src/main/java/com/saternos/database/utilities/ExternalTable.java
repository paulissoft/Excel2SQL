/*
 * Created on Dec 14, 2004
 */
package com.saternos.database.utilities;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import java.text.Normalizer;

/**
 * @author Administrator
 * @version 1.0
 */
public class ExternalTable {
	
    static String newline="\r\n";

    static String double_quote = "\"";
    
    /**
     * Bad File Name for the External Table
     */
    private String badFileName = null;
	
    /**
     * Discard File Name for the External Table
     */
    private String discardFileName = null;
	
    /**
     * List of the External Table columns
     */
    private List<ExternalTableColumn> columns = null;

    /**
     * Comma Separated Values File Extention
     */
    private final String CSV_EXTENTION = ".csv";
	
    /**
     * External table bad file extention
     */
    private final String BAD_EXTENTION = ".bad";
	
    /**
     * External table discard file extention
     */
    private final String DISCARD_EXTENTION = ".dsc";
	
    /**
     * External table log file extention
     */
    private final String LOG_EXTENTION = ".log";
    /**
     * Directory Location referenced by the External Table
     */
    private String dirLocation = null;
	
    /**
     * Log File Name for the External Table
     */	
    private String logFileName = null;
	
    /**
     * Name of the External Table
     */
    private String name = null;
	
    /**
     * Replace all blanks in a given name with underscores
     * @param name
     */

    public ExternalTable(String name){
        setName(name);
        columns = new ArrayList<ExternalTableColumn>();
        badFileName = this.name + BAD_EXTENTION;
        discardFileName = this.name + DISCARD_EXTENTION;
        logFileName = this.name + LOG_EXTENTION;
    }
	
    /**
     * 
     * @param column
     */
    public void addColumn(ExternalTableColumn column)
    {
        columns.add(column);
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
     * Return the SQL String to create the External table
     * which references the newly created CSV files derived
     * from the original Excel sheet
     * @return
     */
    public String getDdl()
    {
        String ddl = "CREATE TABLE " +  getName() +newline+"("+ newline;
		
        Iterator iter = columns.iterator();
        while (iter.hasNext()){
            ExternalTableColumn c = (ExternalTableColumn)iter.next();
            ddl +=c.getColumnDdl();
        }
        ddl = ddl.substring(0, ddl.lastIndexOf(","))+ //remove the last comma
            newline+
            ") "+newline+"ORGANIZATION EXTERNAL"+ newline
            + "(" + newline
            + "  TYPE oracle_loader" + newline
            + "  DEFAULT DIRECTORY load_dir"+ newline
            + "  ACCESS PARAMETERS "+ newline
            + "  ("+ newline
            + "    RECORDS DELIMITED BY NEWLINE"+ newline
            + "    CHARACTERSET WE8MSWIN1252"+ newline
            + "    STRING SIZES ARE IN CHARACTERS"+ newline
            + "    FIELD NAMES ALL FILES IGNORE"+ newline
            + "    BADFILE load_dir:'" + getBadFileName() + "'" + newline
            + "    DISCARDFILE load_dir:'" + getDiscardFileName() + "'" + newline
            + "    LOGFILE load_dir:'" + getLogFileName() + "'" + newline
            + "    FIELDS DATE_FORMAT DATE MASK \"yyyy-mm-dd\" CSV WITHOUT EMBEDDED RECORD TERMINATORS"+ newline
            + "    MISSING FIELD VALUES ARE NULL"+ newline
            + "  )"+ newline
            + "  LOCATION ('"+ getLocation()+"')"+ newline
            +") REJECT LIMIT 0;"+newline+newline+newline;
		
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
     * External table.  This will be the same as the 
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
     * @param columns
     */
    @SuppressWarnings("unchecked")
    public void setColumns(List columns) {
        this.columns = columns;
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
}
