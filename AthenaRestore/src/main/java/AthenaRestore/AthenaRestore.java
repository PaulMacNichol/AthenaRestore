package AthenaRestore;

public class AthenaRestore {
    private static final String PATH = "./xlsx-workbooks/gmacnichol-macros.xlsx";
    private static String       sheetName;

    public static void main ( final String[] args ) {

        System.out.println( "          _   _                      _____           _                 \r\n"
                + "     /\\  | | | |                    |  __ \\         | |                \r\n"
                + "    /  \\ | |_| |__   ___ _ __   __ _| |__) |___  ___| |_ ___  _ __ ___ \r\n"
                + "   / /\\ \\| __| '_ \\ / _ \\ '_ \\ / _` |  _  // _ \\/ __| __/ _ \\| '__/ _ \\\r\n"
                + "  / ____ \\ |_| | | |  __/ | | | (_| | | \\ \\  __/\\__ \\ || (_) | | |  __/\r\n"
                + " /_/    \\_\\__|_| |_|\\___|_| |_|\\__,_|_|  \\_\\___||___/\\__\\___/|_|  \\___|\r\n " );
        System.out.println( "___________________________________________________________________________________" );
        System.out.println(
                "Select and option:\n\t[1] Read from a specific sheet.\n\t[2] Read from entire workbook\n\t[3] Write markup files to XLSX workbook.\n\t[4] Retrieve Text Macros from AthenaHealth\n" );
        System.out.println( ">> " );
        // final TextMacroCollector tmc = new TextMacroCollector();
        // tmc.GET();
        // try {
        // final XLSXReader reader = new XLSXReader( PATH );
        // reader.readAllSheets();
        // }
        // catch ( final IOException e ) {
        // e.printStackTrace();
        // }

    }

}
