/*
 * Created on Dec 14, 2004
 */
package com.saternos.database.utilities;

import org.apache.poi.hssf.usermodel.HSSFCell;

/**
 * @author Casimir Saternos
 * @version 1.0
 */
public class ExternalTableColumn {
	
    static String newline="\r\n";
    
    static String double_quote = "\"";
    
    /**
     * The column length (VARCHAR2)
     */
    private long length = 0;
    /**
     * The column name
     */
    private String name = null;
	
    /**
     * Column precision for numeric types
     */
    private double numericPrecision =0;
	
    /**
     * Space for aligning the outputed SQL
     */
    private final String SPACE="  ";
	
    /**
     * The column type - using "NUMERIC", "STRING" or "DATE"
     */
    private String type = null;
    /**
     * @return
     * 
     * Returns a line of sql according to the following form:
     *      <column_name>    <column_type>,
     */
    public String getColumnDdl() {

        return SPACE + getName()+SPACE +getSqlType()+","+newline;
    }
    /**
     * @return
     * Returns the loader column names according to the form:
     *      <column_name>
     */
    public String getColumnLoaderLine() {
		
        return SPACE + getName() + ","+newline;
    }
	
    /**
     * @return
     */
    public long getLength() {
        return length;
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
     * @return
     */
    private String getSqlType() {

        switch (getType()) {

        case "NUMERIC":
            //handle numeric precision here???
            return "NUMBER";


        case "STRING":
            return "VARCHAR2("+getLength()+")";

        case "DATE":
            return "DATE";
        }
		
		
        return null;
    }

    /**
     * @return
     */
    public String getType() {
        return type;
    }
	
    /**
     * @param length
     */
    public void setLength(long length) {
        this.length = length;
    }
	
    /**
     * @param name
     */
    public void setName(String name) {
        this.name = name;
    }
	
    /**
     * @param numericPrecision
     */
    public void setNumericPrecision(double numericPrecision) {
        this.numericPrecision = numericPrecision;
    }
	
    /**
     * @param type
     */
    public void setType(String type) {
        this.type = type;
    }
}
