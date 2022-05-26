package ExcelUtils;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXWriter {
    private XSSFSheet          activeSheet;
    private final XSSFWorkbook workbook;
    private final File         workbookFile;
    private final String       workbookPath;

    public XLSXWriter ( final String workbookName ) throws InvalidFormatException, IOException {
        workbookPath = "./xlsx-writer/" + workbookName;
        workbookFile = new File( workbookPath );
        workbook = new XSSFWorkbook( workbookFile );
    }

    public void initWorkbook () {
        final File directoryPath = new File( "./output" );
        final File directories[] = directoryPath.listFiles();
        for ( int i = 0; i < directories.length; i++ ) {
            if ( directories[i].isDirectory() ) {
                activeSheet = workbook.createSheet( directories[i].getName() );
            }
        }
    }

    public void initSheet () {

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
