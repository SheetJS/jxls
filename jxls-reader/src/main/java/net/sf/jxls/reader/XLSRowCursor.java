package net.sf.jxls.reader;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Leonid Vysochyn
 */
public interface XLSRowCursor {
    int getCurrentRowNum();
    Row getCurrentRow();
    Sheet getSheet();
    void setSheet(Sheet sheet);
    String getSheetName();
    void setSheetName(String sheetName);
    Row next();
    boolean hasNext();
    void reset();
    void setCurrentRowNum(int rowNum);
    void moveForward();
    void moveBackward();
}
