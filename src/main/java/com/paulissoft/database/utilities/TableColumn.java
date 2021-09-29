package com.paulissoft.database.utilities;

import java.text.Normalizer;

/**
 * @author Casimir Saternos
 * @version 1.0
 */
public class TableColumn {

    /**
     * The maximum column length for strings in characters (not bytes).
     */
    private int stringLength = 0; // at least

    /**
     * The column name.
     */
    private String name = null;

    /**
     * Column precision for numeric types.
     */
    private int numericPrecision = 0;

    /**
     * The maximum column length for numbers in characters.
     */
    private int numericLength = 0;

    /**
     * The maximum column length for dates in characters.
     */
    private int dateLength = 0;

    /**
     * Spaces for aligning the outputted SQL.
     */
    static final String SPACES = "  ";

    /**
     * The application run-time settings.
     */
    private Settings settings = null;

    /**
     * The constructor.
     *
     * @param pSettings  The application run-time settings
     */
    public TableColumn(final Settings pSettings) {
        this.settings = pSettings;
    }

    /**
     * @return
     *
     * Returns a line of sql according to the following form:
     *      <column_name>    <column_type>,
     */
    public String getColumnDdl() {

        return TableColumn.SPACES + getName() + TableColumn.SPACES
            + getSqlType() + "," + Settings.NL;
    }
    /**
     * @return
     * Returns the loader column names according to the form:
     *      <column_name>
     */
    public String getColumnLoaderLine() {

        return TableColumn.SPACES + getName() + "," + Settings.NL;
    }

    /**
     * @return The maximum length for a value converted to a string.
     */
    public int getStringLength() {
        return stringLength;
    }

    /**
     * @return The (quoted) name of the column.
     */
    public String getName() {
        return (name == null ? null : Settings.QQ + name + Settings.QQ);
    }

    /**
     * @return The precision of a numeric value.
     */
    public int getNumericPrecision() {
        return numericPrecision;
    }

    /**
     * @return The numeric length.
     */
    public int getNumericLength() {
        return numericLength;
    }

    /**
     * @return The date length.
     */
    public int getDateLength() {
        return dateLength;
    }

    /**
     * @return The SQL type (NUMERIC, DATE or STRING).
     */
    private String getSqlType() {
        final String stringDatatype =
            (settings.sqlDatabase.equals(Settings.ORACLE)
             ? "VARCHAR2"
             : "VARCHAR");

        Integer stringColumnSize = settings.stringColumnSize;

        // Use SQL ANSI datatypes
        if (stringLength == 0 && numericLength == 0 && dateLength == 0) {
            if (stringColumnSize == null) {
                stringColumnSize = 1;
            }
            return stringDatatype + "(" + stringColumnSize + ")";
        } else if (stringLength > 0
                   || (numericLength > 0 && dateLength > 0)) {
            // column includes a non empty string value or
            // both numeric and date values: convert it to VARCHAR2
            if (stringColumnSize == null) {
                stringColumnSize = Math.max(Math.max(stringLength, numericLength), dateLength);
            }
            return stringDatatype + "(" + stringColumnSize + ")";
        } else if (numericLength > 0) {
            // handle numeric precision here???
            return "DECIMAL";
        } else {
            return "DATE";
        }
    }

    /**
     * @param pStringLength  The string length.
     */
    public void setStringLength(final int pStringLength) {
        if (this.stringLength < pStringLength) {
            this.stringLength = pStringLength;
        }
    }

    /**
     * @param pName  The column name.
     */
    public void setName(final String pName) {
        this.name =
            Normalizer.normalize(pName, Normalizer.Form.NFD)
            .replaceAll("\\p{M}", "").trim();
    }

    /**
     * @param pNumericPrecision  The numeric precision to set
     *                           if longer than the actual
     */
    public void setNumericPrecision(final int pNumericPrecision) {
        this.numericPrecision = pNumericPrecision;
    }

    /**
     * @param pNumericLength  The numeric length to set
     *                        if longer than the actual
     */
    public void setNumericLength(final int pNumericLength) {
        if (this.numericLength < pNumericLength) {
            this.numericLength = pNumericLength;
        }
    }

    /**
     * @param pDateLength  The date length to set if longer than the actual
     */
    public void setDateLength(final int pDateLength) {
        if (this.dateLength < pDateLength) {
            this.dateLength = pDateLength;
        }
    }
}
