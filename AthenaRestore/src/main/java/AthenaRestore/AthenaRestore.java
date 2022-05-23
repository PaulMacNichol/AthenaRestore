package AthenaRestore;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import AthenaModels.TextMacro;
import ExcelUtils.XLSXReader;

public class AthenaRestore {
    private static final String PATH      = "./data/gmacnichol-macros.xlsx";
    private static final String HPI_SHEET = "HPI";

    public static void main ( final String[] args ) {
        List<TextMacro> macroData = null;
        try {
            final XLSXReader reader = new XLSXReader( PATH, HPI_SHEET );
            reader.readMacros();
            macroData = reader.getListOfMacros();
        }
        catch ( final IOException e ) {
            e.printStackTrace();
        }

        final Iterator<TextMacro> macroIterator = macroData.iterator();
        while ( macroIterator.hasNext() ) {
            final TextMacro currentMacro = macroIterator.next();
            final String fileName = currentMacro.getShortcut() + ".html";
            final String pathName = "./output/" + fileName;
            final File xmlMacro = new File( pathName );

            try {
                xmlMacro.createNewFile();
            }
            catch ( final IOException e ) {
                System.out.println( "Error creating file." );
                e.printStackTrace();
            }

            try {
                final FileWriter macroWriter = new FileWriter( pathName );
                macroWriter.write( currentMacro.getExpansion() );
                macroWriter.close();
            }
            catch ( final IOException e ) {
                System.out.print( "Error writing file." );
                e.printStackTrace();
            }
        }
    }

}
