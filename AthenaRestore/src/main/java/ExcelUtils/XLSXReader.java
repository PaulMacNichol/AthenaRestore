package ExcelUtils;

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
    private XSSFSheet       sheet;
    private XSSFWorkbook    workbook;
    private List<TextMacro> listOfMacros;

    public XLSXReader () {
    }

    public XLSXReader ( final String pathName, final String sheetName ) throws IOException {
        setPathName( pathName );
        setSheetName( sheetName );
        workbook = new XSSFWorkbook( getPathName() );
        sheet = workbook.getSheet( getSheetName() );
        listOfMacros = new ArrayList<TextMacro>();
    }

    public void readMacros () {
        final Iterator<Row> rowIterator = sheet.rowIterator();
        while ( rowIterator.hasNext() ) {
            rowIterator.next();
            final Row currentRow = rowIterator.next();
            final Iterator<Cell> cellIterator = currentRow.cellIterator();
            while ( cellIterator.hasNext() ) {

                final Cell shortcut = cellIterator.next();
                final Cell expansion = cellIterator.next();
                if ( nullMacroCheck( shortcut.getStringCellValue(), expansion.getStringCellValue() ) ) {
                    final TextMacro tm = new TextMacro( shortcut.getStringCellValue(), null, null,
                            expansion.getStringCellValue() );
                    listOfMacros.add( tm );
                }
            }
        }
    }

    private boolean nullMacroCheck ( final String shortcut, final String expansion ) {
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
        return sheet;
    }

    /**
     * @param sheet
     *            the sheet to set
     */
    public void setSheet ( final XSSFSheet sheet ) {
        this.sheet = sheet;
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
