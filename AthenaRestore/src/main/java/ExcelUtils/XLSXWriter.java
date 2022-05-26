package ExcelUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXWriter {
    private XSSFSheet          activeSheet;
    private final XSSFWorkbook workbook;
    private final File         workbookFile;
    private final String       workbookPath;

    public XLSXWriter () {
        this.workbookFile = null;
        this.workbookPath = "";
        this.workbook = new XSSFWorkbook();
    }

    // public XLSXWriter ( final String workbookName ) throws
    // InvalidFormatException, IOException {
    // workbookPath = "./xlsx-workbooks/" + workbookName + ".xlsx";
    // workbookFile = new File( workbookPath );
    // workbookFile.createNewFile();
    // workbook = new XSSFWorkbook( workbookFile );
    // }

    public void initWorkbook () throws IOException {
        final File directoryPath = new File( "./text-macros/" );
        final File directories[] = directoryPath.listFiles();
        for ( int i = 0; i < directories.length; i++ ) {
            if ( directories[i].isDirectory() ) {
                setActiveSheet( workbook.createSheet( directories[i].getName() ) );
                initSheet( directories[i] );
            }
        }
    }

    public void initSheet ( final File currentDirectory ) throws IOException {
        final File[] files = currentDirectory.listFiles();
        for ( int i = 0; i < files.length; i++ ) {
            final Row row = activeSheet.createRow( i );
            if ( row.getRowNum() == 0 ) {
                final Cell sh = row.createCell( 0 );
                sh.setCellValue( "Shortcut" );
                final Cell exp = row.createCell( 1 );
                sh.setCellValue( "Expansion" );

            }
            else {
                final Cell shortcut = row.createCell( 0 );
                shortcut.setCellValue( files[i].getName() );
                final Cell expansion = row.createCell( 1 );
                final String data = FileUtils.readFileToString( files[i], "UTF-8" );
                expansion.setCellValue( data );
            }
        }

    }

    public void saveWorksheet ( final String workSheetName ) throws IOException {
        final FileOutputStream out = new FileOutputStream( new File( "./xlsx-workbooks/" + workSheetName + ".xlsx" ) );
        workbook.write( out );
        out.close();
    }

    /**
     * @return the activeSheet
     */
    public XSSFSheet getActiveSheet () {
        return activeSheet;
    }

    /**
     * @param activeSheet
     *            the activeSheet to set
     */
    public void setActiveSheet ( final XSSFSheet activeSheet ) {
        this.activeSheet = activeSheet;
    }

    /**
     * @return the workbook
     */
    public XSSFWorkbook getWorkbook () {
        return workbook;
    }
}
