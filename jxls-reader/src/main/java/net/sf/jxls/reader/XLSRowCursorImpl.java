package net.sf.jxls.reader;

import java.util.NoSuchElementException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * @author Leonid Vysochyn
 */
public class XLSRowCursorImpl implements XLSRowCursor {
    int currentRowNum;
    HSSFSheet sheet;
    String sheetName;


    public XLSRowCursorImpl(HSSFSheet sheet) {
        this.sheet = sheet;
    }


    public XLSRowCursorImpl(String sheetName, HSSFSheet sheet) {
        this.sheetName = sheetName;
        this.sheet = sheet;
    }

    public int getCurrentRowNum() {
        return currentRowNum;
    }

    public HSSFRow getCurrentRow() {
        return sheet.getRow( currentRowNum );
    }


    public HSSFSheet getSheet() {
        return sheet;
    }

    public void setSheet(HSSFSheet sheet) {
        this.sheet = sheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public HSSFRow next() {
        if( hasNext() ){
            return sheet.getRow( currentRowNum++ );
        }
        throw new NoSuchElementException();
    }

    public boolean hasNext() {
        return (currentRowNum <= sheet.getLastRowNum());
    }

    public void reset() {
        currentRowNum = 0;
    }

    public void setCurrentRowNum(int rowNum) {
        currentRowNum = rowNum;
    }

    public void moveForward() {
        currentRowNum++;
    }

    public void moveBackward() {
        currentRowNum--;
    }
}
