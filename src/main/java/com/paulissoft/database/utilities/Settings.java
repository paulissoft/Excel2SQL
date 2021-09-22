package com.paulissoft.database.utilities;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.beust.jcommander.IParameterValidator;
import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;
import com.beust.jcommander.ParameterException;

class Settings {

    final static String ORACLE = "Oracle";
    final static String POSTGRESQL = "PostgresQL";

    @Parameter(names = "--verbose", description = "Level of verbosity")
    protected Integer verbose = 1;

    @Parameter(names = "--sheet-name-expression", description = "Sheet name(s) must match this regular expression")
    protected String sheetNameExpression = ".*";

    @Parameter(names = "--sql-table-name", description = "A list of SQL table name(s) to use instead of the sheet name(s)")
    protected List<String> tableNames = new ArrayList<>();

    @Parameter(names = "--column-separator", description = "The column separator")
    protected String columnSeparator = ",";

    @Parameter(names = "--enclosure-string", description = "The enclosure string")
    protected String enclosureString = "\"";

    @Parameter(names = "--encoding", description = "The encoding to use (default \"windows-1252\")")
    protected String encoding = "windows-1252";
    
    @Parameter(names = "--write-bom", description = "Write the BOM at the beginning of the file.")
    protected boolean writeBOM = false;

    @Parameter(names = "--no-header", description = "The first row does NOT contain the column names")
    protected boolean noHeader = false;

    @Parameter(names = "--sql-database", description = "The SQL database (Oracle, PostgresQL)", required = false, validateWith = ValidSqlDatabases.class)
    protected String sqlDatabase = "Oracle";

    @Parameter(names = "--one-table", description = "Each sheet to one table?")
    protected boolean oneTable = false;

    @Parameter(names = "--add-metadata", description = "Add metadata like sheet name and row number (starting from 1) in the CSV?")
    protected boolean addMetadata = false;

    @Parameter(names = { "--help", "-h" }, description = "This help", help = true)
    protected boolean help;
    
    /**
     * The Excel Spreadsheets (.xls or .xlsx) that are being accessed
     */
    @Parameter(description = "spreadsheet...", required = true)
    protected List<String> spreadsheets = new ArrayList<String>();

    public static class ValidSqlDatabases implements IParameterValidator {
        @Override
        public void validate(String name, String value) throws ParameterException {
            List<String> databases = Arrays.asList(ORACLE, POSTGRESQL);
        
            if (!databases.contains(value)) {
                throw new ParameterException("Parameter " + name + " (" + value +") is not a valid database (" + databases + ")");
            }
        }
    }
}