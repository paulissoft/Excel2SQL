/*
 * Created on Dec 14, 2004
 */
package com.saternos.database.utilities;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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
     * List of the External Table columns
     */
    private List<ExternalTableColumn> columns = null;

    /**
     * Comma Separated Values File Extention
     */
    private final String CSV_EXTENTION= ".csv";
	
    /**
     * External table bad file extention
     */
    private final String BAD_EXTENTION= ".bad";
	
    /**
     * External table log file extention
     */
    private final String LOG_EXTENTION= ".log";
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
        this.name=name.replaceAll(" ","_");
        columns = new ArrayList<ExternalTableColumn>();
        badFileName=this.name + BAD_EXTENTION;
        logFileName=this.name + LOG_EXTENTION;
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
        String ddl = "CREATE TABLE " +  getName() +newline+" ("+ newline;
		
        Iterator iter = columns.iterator();
        while (iter.hasNext()){
            ExternalTableColumn c = (ExternalTableColumn)iter.next();
            ddl +=c.getColumnDdl();
        }
        ddl = ddl.substring(0, ddl.lastIndexOf(","))+ //remove the last comma
            newline+
            ") "+newline+"ORGANIZATION EXTERNAL"+ newline
            + "   (" + newline
            + "         TYPE oracle_loader" + newline
            + "         DEFAULT DIRECTORY load_dir"+ newline
            + "         ACCESS PARAMETERS "+ newline
            + "         ("+ newline
            + "               RECORDS DELIMITED BY NEWLINE"+ newline
            + "               FIELD NAMES ALL FILES"+ newline
            + "               BADFILE load_dir:'" + getBadFileName() +"'"+ newline
            + "               LOGFILE load_dir:'" + getLogFileName() +"'"+ newline
            // + "               FIELDS TERMINATED BY ','"+newline
            + "               FIELDS DATE_FORMAT DATE MASK \"yyyy-mm-dd\" CSV WITHOUT EMBEDDED RECORD TERMINATORS"+ newline
            + "               MISSING FIELD VALUES ARE NULL"+ newline
            /*
              + "               ( "+ newline
            */
            ;
        /* not needed due to FIELD NAMES ALL FILES */
        /*
          iter = columns.iterator();
          while (iter.hasNext()){
          ExternalTableColumn c = (ExternalTableColumn)iter.next();
          ddl +="                 "+c.getColumnLoaderLine();
          }
          ddl = ddl.substring(0, ddl.lastIndexOf(","))
          + newline
          + "         )";
        */
        ddl +=
            "               )"+ newline
            + "             LOCATION ('"+ getLocation()+"')"+ newline
            +") REJECT LIMIT UNLIMITED;"+newline+newline+newline;
		
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
        this.name = name;
    }
}
