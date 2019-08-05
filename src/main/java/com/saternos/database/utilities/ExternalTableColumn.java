/*
 * Created on Dec 14, 2004
 */
package com.saternos.database.utilities;

import java.text.Normalizer;

/**
 * @author Casimir Saternos
 * @version 1.0
 */
public class ExternalTableColumn {
	
    static String newline="\r\n";
    
    static String double_quote = "\"";
    
    /**
     * The maximum column length for strings in characters (not bytes)
     */
    private long stringLength = 0; // at least
    /**
     * The column name
     */
    private String name = null;
	
    /**
     * Column precision for numeric types
     */
    private double numericPrecision = 0;
    
    /**
     * The maximum column length for numbers in characters
     */
    private long numericLength = 0;
	
    /**
     * The maximum column length for dates in characters
     */
    private long dateLength = 0;
	
    /**
     * Spaces for aligning the outputted SQL
     */
    private final String SPACES = "  ";
	
    /**
     * @return
     * 
     * Returns a line of sql according to the following form:
     *      <column_name>    <column_type>,
     */
    public String getColumnDdl() {

        return SPACES + getName() + SPACES + getSqlType() + "," + newline;
    }
    /**
     * @return
     * Returns the loader column names according to the form:
     *      <column_name>
     */
    public String getColumnLoaderLine() {
		
        return SPACES + getName() + "," + newline;
    }
	
    /**
     * @return
     */
    public long getStringLength() {
        return stringLength;
    }

    /**
     * @return
     */
    public String getName() {
        return (name == null ? null : double_quote + name + double_quote);
    }
	
    /**
     * @return double
     */
    public double getNumericPrecision() {
        return numericPrecision;
    }
    
    /**
     * @return long
     */
    public long getNumericLength() {
        return numericLength;
    }
    
    /**
     * @return long
     */
    public long getDateLength() {
        return dateLength;
    }
    
    /**
     * @return
     */
    private String getSqlType() {

        if (stringLength == 0 && numericLength == 0 && dateLength == 0) {
            return "VARCHAR2(1 CHAR)";
        } else if (stringLength > 0 ||
                   (numericLength > 0 && dateLength > 0)) {
            // column includes a non empty string value or both numeric and date values: convert it to VARCHAR2
            return "VARCHAR2(" + Math.max(Math.max(stringLength, numericLength), dateLength) + " CHAR)";
        } else if (numericLength > 0) {
            // handle numeric precision here???
            return "NUMBER";
        } else {
            return "DATE";
        }
    }

    /**
     * @param stringLength
     */
    public void setStringLength(long stringLength) {
        if (this.stringLength < stringLength)
            this.stringLength = stringLength;
    }
	
    /**
     * @param name
     */
    public void setName(String name) {
        this.name = Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
    }
	
    /**
     * @param numericPrecision
     */
    public void setNumericPrecision(double numericPrecision) {
        this.numericPrecision = numericPrecision;
    }
	
    /**
     * @param numericLength
     */
    public void setNumericLength(long numericLength) {
        if (this.numericLength < numericLength)
            this.numericLength = numericLength;
    }

    /**
     * @param dateLength
     */
    public void setDateLength(long dateLength) {
        if (this.dateLength < dateLength)
            this.dateLength = dateLength;
    }
	
}
