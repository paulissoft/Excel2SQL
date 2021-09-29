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
     * Bad File Name for the Table.
     */
    private String badFileName = null;

    /**
     * Discard File Name for the Table.
     */
    private String discardFileName = null;

    /**
     * List of the Table columns.
     */
    private List<TableColumn> columns = null;

    /**
     * Comma Separated Values File Extension.
     */
    static final String CSV_EXTENSION = ".csv";

    /**
     * table bad file extension.
     */
    static final String BAD_EXTENSION = ".bad";

    /**
     * table discard file extension.
     */
    static final String DISCARD_EXTENSION = ".dsc";

    /**
     * table log file extension.
     */
    static final String LOG_EXTENSION = ".log";

    /**
     * Directory Location referenced by the Table.
     */
    private String dirLocation = null;

    /**
     * Log File Name for the Table.
     */
    private String logFileName = null;

    /**
     * Name of the Table.
     */
    private String name = null;

    /** The run-time settings. */
    private Settings settings = null;

    /**
     * The Table constructor.
     *
     * @param pName      The table name
     * @param pSettings  The run-time settings
     */
    public Table(final String pName, final Settings pSettings) {
        setName(pName);
        this.settings = pSettings;
        columns = new ArrayList<TableColumn>();
        badFileName = this.name + Table.BAD_EXTENSION;
        discardFileName = this.name + Table.DISCARD_EXTENSION;
        logFileName = this.name + Table.LOG_EXTENSION;
    }

    private String getOracleCharacterSet() {
        switch (getEncoding()) {
        case "windows-1252":
            return "WE8MSWIN1252";
        case "UTF-8":
            return "AL32UTF8";
        default:
            throw new RuntimeException("Should have encoding of \"windows-1252\" or \"UTF-8\"");
        }
    }

    /**
     * Add a olumn.
     *
     * @param column
     */
    public void addColumn(final TableColumn column) {
        columns.add(column);
    }

    /**
     * Add a first column.
     *
     * @param column
     */
    public void addColumnFirst(final TableColumn column) {
        columns.add(0, column);
    }

    /**
     * Get the column at index i (at least 0).
     *
     * @param index  The index to get
     *
     * @return A column
     */
    public TableColumn getColumn(final int index) {
        return columns.get(index);
    }

    /**
     * Get the number of columns.
     *
     * @return columns
     */
    public int getNrColumns() {
        return columns.size();
    }

    /**
     * Get the bad file name.
     *
     * @return the bad file name
     */
    public String getBadFileName() {
        return badFileName;
    }

    /**
     * Get the discard file name.
     *
     * @return the discard file name
     */
    public String getDiscardFileName() {
        return discardFileName;
    }

    /**
     * Get the columns.
     *
     * @return The columns
     */
    public List getColumns() {
        return columns;
    }

    /**
     * Return the SQL String to create the table
     * which references the newly created CSV files derived
     * from the original Excel sheet.
     *
     * @return DDL text
     */
    public String getDdl() {
        String ddl =
            "CREATE TABLE "
            + (settings.sqlDatabase.equals(Settings.POSTGRESQL) ? "IF NOT EXISTS " : "")
            + getName() + Settings.NL + "(" + Settings.NL;
        Iterator iter = columns.iterator();

        while (iter.hasNext()) {
            TableColumn c = (TableColumn) iter.next();
            ddl += c.getColumnDdl();
        }
        ddl = ddl.substring(0, ddl.lastIndexOf(","))
            //remove the last comma
            + Settings.NL
            + ")";

        if (settings.sqlDatabase.equals(Settings.ORACLE)) {
            ddl += Settings.NL
                + "ORGANIZATION EXTERNAL" + Settings.NL
                + "(" + Settings.NL
                + "  TYPE oracle_loader" + Settings.NL
                + "  DEFAULT DIRECTORY load_dir" + Settings.NL
                + "  ACCESS PARAMETERS " + Settings.NL
                + "  (" + Settings.NL
                + "    RECORDS DELIMITED BY SETTINGS.NL" + Settings.NL
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
                + "    STRING SIZES ARE IN BYTES" + Settings.NL
                + "    FIELD NAMES ALL FILES IGNORE" + Settings.NL
                + "    BADFILE load_dir:'" + getBadFileName() + "'" + Settings.NL
                + "    DISCARDFILE load_dir:'" + getDiscardFileName() + "'" + Settings.NL
                + "    LOGFILE load_dir:'" + getLogFileName() + "'" + Settings.NL
                // Oracle 12C syntax
                + "    FIELDS TERMINATED BY '" + getFieldSeparator()
                + "' OPTIONALLY ENCLOSED BY '" + getEnclosureString()
                + "' DATE_FORMAT DATE MASK \"yyyy-mm-dd\"" + Settings.NL
                // Oracle 12CR2 syntax
                // + "    FIELDS DATE_FORMAT DATE MASK \"yyyy-mm-dd\" CSV WITHOUT EMBEDDED RECORD TERMINATORS" + Settings.NL
                + "    MISSING FIELD VALUES ARE NULL" + Settings.NL
                + "  )" + Settings.NL
                + "  LOCATION ('" + getLocation() + "')" + Settings.NL
                + ") REJECT LIMIT 0" + Settings.NL + Settings.NL + Settings.NL;
        }

        ddl += ";" + Settings.NL;

        return ddl;
    }

    /**
     * Get the directory location.
     *
     * @return The directory location
     */
    public String getDirLocation() {
        return dirLocation;
    }

    /**
     * Get the location of the file referenced by the
     * table.  This will be the same as the
     * table name with a .csv extension.
     *
     * @return The data file location.
     */
    public String getLocation() {
        return name + Table.CSV_EXTENSION;
    }

    /**
     * Get the log file.
     *
     * @return The log file
     */
    public String getLogFileName() {
        return logFileName;
    }

    /**
     * Get the table name.
     *
     * @return The table name
     */
    public String getName() {
        return (name == null ? null : Settings.QQ + name + Settings.QQ);
    }

    /**
     * Set the bad file name.
     *
     * @param pBadFileName  the bad file name
     */
    public void setBadFileName(final String pBadFileName) {
        this.badFileName = pBadFileName;
    }

    /**
     * Set the discard file name.
     *
     * @param pDiscardFileName  the discard file name
     */
    public void setDiscardFileName(final String pDiscardFileName) {
        this.discardFileName = pDiscardFileName;
    }

    /**
     * Set the directory location.
     *
     * @param pDirLocation  The directory location
     */
    public void setDirLocation(final String pDirLocation) {
        this.dirLocation = pDirLocation;
    }

    /**
     * Set the log file name.
     *
     * @param pLogFileName  The log file name
     */
    public void setLogFileName(final String pLogFileName) {
        this.logFileName = pLogFileName;
    }

    /**
     * Set the table name.
     *
     * @param pName  The table name
     */
    public void setName(final String pName) {
        this.name = Normalizer.normalize(pName, Normalizer.Form.NFD).replaceAll("\\p{M}", "").trim();
    }

    /**
     * Get the normalized table name.
     *
     * @param name  The table name
     *
     * @return  The normalized table name
     */
    public static String getName(final String name) {
        if (name == null) {
            return null;
        }

        return
            Settings.QQ
            + Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "").trim()
            + Settings.QQ;
    }

    /**
     * Get the field separator.
     *
     * @return The field separator
     */
    public String getFieldSeparator() {
        return this.settings.columnSeparator;
    }

    /**
     * Get the enclosure string.
     *
     * @return The enclosure string
     */
    public String getEnclosureString() {
        return this.settings.enclosureString;
    }

    /**
     * Get the encoding.
     *
     * @return The encoding
     */
    public String getEncoding() {
        return this.settings.encoding;
    }

    /**
     * Get the preamble for the DDL file.
     *
     * @param settings  The run-time settings
     *
     * @return the preamble
     */
    protected static String preamble(final Settings settings) {
        if (settings.sqlDatabase.equals(Settings.ORACLE)) {
            return
                "CREATE /*OR REPLACE*/ DIRECTORY load_dir AS '" + Settings.PWD + "'" + Settings.NL
                + ";" + Settings.NL + Settings.NL;
        } else {
            return "";
        }
    }
}
