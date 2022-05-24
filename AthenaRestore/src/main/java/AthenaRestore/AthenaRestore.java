package AthenaRestore;

import java.io.IOException;

import AthenaCom.TextMacroCollector;
import ExcelUtils.XLSXReader;

public class AthenaRestore {
    private static final String PATH      = "./data/gmacnichol-macros.xlsx";
    private static final String HPI_SHEET = "HPI";

    public static void main ( final String[] args ) {
        final TextMacroCollector tmc = new TextMacroCollector();
        tmc.GET();
        try {
            final XLSXReader reader = new XLSXReader( PATH );
            reader.readAllSheets();
        }
        catch ( final IOException e ) {
            e.printStackTrace();
        }

    }

}
