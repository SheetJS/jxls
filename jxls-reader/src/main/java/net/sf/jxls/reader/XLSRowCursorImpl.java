package net.sf.jxls.reader;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Leonid Vysochyn
 */
public class XLSRowCursorImpl implements XLSRowCursor {
    int currentRowNum;
    Sheet sheet;
    String sheetName;


    public XLSRowCursorImpl(Sheet sheet) {
        this.sheet = sheet;
    }


    public XLSRowCursorImpl(String sheetName, Sheet sheet) {
        this.sheetName = sheetName;
        this.sheet = sheet;
    }

    public int getCurrentRowNum() {
        return currentRowNum;
    }

    public Row getCurrentRow() {
        return sheet.getRow( currentRowNum );
    }


    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Row next() {
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
