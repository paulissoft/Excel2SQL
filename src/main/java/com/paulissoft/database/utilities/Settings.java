package com.paulissoft.database.utilities;

import java.io.File;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.beust.jcommander.IParameterValidator;
import com.beust.jcommander.Parameter;
import com.beust.jcommander.ParameterException;

class Settings {
    /** The newline string. */
    static final String NL = "\r\n";
    /** The double quote. */
    static final String QQ = "\"";
    /**
     * Present working directory.
     */
    static final String PWD = new File("").getAbsolutePath();

    /** The Oracle database. */
    static final String ORACLE = "Oracle";

    /** The PostgreSQL database. */
    static final String POSTGRESQL = "PostgresQL";

    /** The verbosity. */
    @Parameter(names = "--verbose", description = "Level of verbosity")
    protected Integer verbose = 1;

    /** The sheet name expression. */
    @Parameter(names = "--sheet-name-expression",
               description = "Sheet name(s) must match this regular expression")
    protected String sheetNameExpression = ".*";

    /** A list of SQL table names to be used instead of the sheet names. */
    @Parameter(names = "--sql-table-names",
               description = "A list of SQL table name(s) to use instead of the sheet name(s)")
    protected List<String> tableNames = new ArrayList<>();

    /** The CSV column separator. */
    @Parameter(names = "--column-separator",
               description = "The column separator")
    protected String columnSeparator = ",";

    /** The CSV enclosure string. */
    @Parameter(names = "--enclosure-string",
               description = "The enclosure string")
    protected String enclosureString = "\"";

    /** The encoding to use which defaults to windows-1252 the standard for CSV. */
    @Parameter(names = "--encoding",
               description = "The encoding to use (default \"windows-1252\")")
    protected String encoding = "windows-1252";

    /** Write the BOM at the beginning of the file? */
    @Parameter(names = "--write-bom",
               description = "Write the BOM at the beginning of the file.")
    protected boolean writeBOM = false;

    /** The first header row (0 means no header). */
    @Parameter(names = "--header-row-from",
               description = "The first header row (0 means no header)")
    protected Integer headerRowFrom = 1;

    /** The last header row (0 means no header). */
    @Parameter(names = "--header-row-till",
               description = "The last header row (0 means no header)")
    protected Integer headerRowTill = 1;

    /** The SQL database to write code for. */
    @Parameter(names = "--sql-database",
               description = "The SQL database (Oracle, PostgresQL)",
               required = false,
               validateWith = ValidSqlDatabases.class)
    protected String sqlDatabase = "Oracle";

    /** Add metadata like sheet name and row number (starting from 1) in the CSV? */
    @Parameter(names = "--add-metadata",
               description = "Add metadata like sheet name and row number (starting from 1) in the CSV?")
    protected boolean addMetadata = false;

    /** Help? */
    @Parameter(names = { "--help", "-h" }, description = "This help", help = true)
    protected boolean help;

    /** The sql string size. */
    @Parameter(names = { "--string-column-size"},
               description = "Use this for the string column size and not the maximum Excel data size.")
    protected Integer stringColumnSize = null;

    /**
     * The Excel Spreadsheet (.xls or .xlsx) that is being accessed.
     */
    @Parameter(description = "spreadsheet", required = true)
    protected String spreadsheet = null;

    public static class ValidSqlDatabases implements IParameterValidator {
        @Override
        public void validate(final String name, final String value) throws ParameterException {
            List<String> databases = Arrays.asList(ORACLE, POSTGRESQL);

            if (!databases.contains(value)) {
                throw new ParameterException("Parameter " + name + " (" + value
                                             + ") is not a valid database (" + databases + ")");
            }
        }
    }
}
