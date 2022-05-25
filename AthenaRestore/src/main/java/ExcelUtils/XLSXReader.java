package ExcelUtils;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import AthenaModels.TextMacro;

public class XLSXReader {
    private String          pathName;
    private String          sheetName;
    private XSSFSheet       activeSheet;
    private XSSFWorkbook    workbook;
    private List<TextMacro> listOfMacros;

    public XLSXReader () {
    }

    public XLSXReader ( final String pathName ) throws IOException {
        setPathName( pathName );
        setSheetName( null );
        workbook = new XSSFWorkbook( getPathName() );
        listOfMacros = new ArrayList<TextMacro>();
    }

    public XLSXReader ( final String pathName, final String sheetName ) throws IOException {
        setPathName( pathName );
        setSheetName( sheetName );
        workbook = new XSSFWorkbook( getPathName() );
        activeSheet = workbook.getSheet( getSheetName() );
        listOfMacros = new ArrayList<TextMacro>();
    }

    public void readAllSheets () {
        final int numSheets = workbook.getNumberOfSheets();
        for ( int i = 0; i < numSheets; i++ ) {
            System.out.println( "Working on " + workbook.getSheetName( i ) );
            if ( !workbook.getSheetName( i ).equals( "Encounter Plans" ) ) {
                setSheetName( workbook.getSheetName( i ) );

                setActiveSheetByName( getSheetName() );
                readMacros();
                generateMarcosAsHTML();
                listOfMacros.clear();
            }
        }

    }

    public void readMacros () {
        final Iterator<Row> rowIterator = activeSheet.rowIterator();
        rowIterator.next();
        while ( rowIterator.hasNext() ) {
            final Row currentRow = rowIterator.next();
            final Iterator<Cell> cellIterator = currentRow.cellIterator();
            while ( cellIterator.hasNext() ) {

                final Cell shortcut = cellIterator.next();
                Cell expansion = null;
                if ( cellIterator.hasNext() ) {
                    expansion = cellIterator.next();
                }
                if ( ! ( shortcut == null || expansion == null ) ) {
                    if ( nullMacroCheck( shortcut.getStringCellValue(), expansion.getStringCellValue() ) ) {
                        final TextMacro tm = new TextMacro( shortcut.getStringCellValue(), null, null,
                                expansion.getStringCellValue() );
                        listOfMacros.add( tm );
                    }
                }
            }
        }
    }

    public void generateMarcosAsHTML () {

        final Iterator<TextMacro> macroIterator = getListOfMacros().iterator();
        while ( macroIterator.hasNext() ) {
            final TextMacro currentMacro = macroIterator.next();
            final String fileName = currentMacro.getShortcut() + ".html";
            final String pathName = "./output/" + sheetName;
            final File folder = new File( pathName );
            if ( !folder.exists() ) {
                folder.mkdirs();
            }
            final File xmlMacro = new File( pathName + "/" + fileName );
            if ( !xmlMacro.exists() ) {
                try {
                    xmlMacro.createNewFile();
                }
                catch ( final IOException e ) {
                    System.out.println( "Error creating file." );
                    e.printStackTrace();
                }
                try {
                    final FileWriter macroWriter = new FileWriter( pathName + "/" + fileName );
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

    private boolean nullMacroCheck ( final String shortcut, final String expansion ) {
        if ( shortcut == null || expansion == null ) {
            return false;
        }
        if ( shortcut.isBlank() ) {
            return false;
        }

        if ( shortcut.isEmpty() ) {
            return false;
        }

        if ( expansion.isBlank() ) {
            return false;
        }

        if ( expansion.isEmpty() ) {
            return false;
        }
        return true;
    }

    /**
     * @return the listOfMacros
     */
    public List<TextMacro> getListOfMacros () {
        return listOfMacros;
    }

    /**
     * @param listOfMacros
     *            the listOfMacros to set
     */
    public void setListOfMacros ( final List<TextMacro> listOfMacros ) {
        this.listOfMacros = listOfMacros;
    }

    /**
     * @return the pathName
     */
    public String getPathName () {
        return pathName;
    }

    /**
     * @param pathName
     *            the pathName to set
     */
    public void setPathName ( final String pathName ) {
        this.pathName = pathName;
    }

    /**
     * @return the sheetName
     */
    public String getSheetName () {
        return sheetName;
    }

    /**
     * @param sheetName
     *            the sheetName to set
     */
    public void setSheetName ( final String sheetName ) {
        this.sheetName = sheetName;
    }

    /**
     * @return the sheet
     */
    public XSSFSheet getSheet () {
        return activeSheet;
    }

    /**
     * @param sheet
     *            the sheet to set
     */
    public void setSheet ( final XSSFSheet sheet ) {
        this.activeSheet = sheet;
    }

    public void setActiveSheetByName ( final String name ) {
        this.activeSheet = workbook.getSheet( name );
    }

    /**
     * @return the workbook
     */
    public XSSFWorkbook getWorkbook () {
        return workbook;
    }

    /**
     * @param workbook
     *            the workbook to set
     */
    public void setWorkbook ( final XSSFWorkbook workbook ) {
        this.workbook = workbook;
    }
}
