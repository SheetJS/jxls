package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * @author Leonid Vysochyn
 */
public interface XLSRowCursor {
    int getCurrentRowNum();
    HSSFRow getCurrentRow();
    HSSFSheet getSheet();
    void setSheet(HSSFSheet sheet);
    HSSFRow next();
    boolean hasNext();
    void reset();
    void setCurrentRowNum(int rowNum);
    void moveForward();
    void moveBackward();
}
