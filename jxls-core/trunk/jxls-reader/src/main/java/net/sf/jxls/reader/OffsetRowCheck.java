package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFRow;

/**
 * @author Leonid Vysochyn
 */
public interface OffsetRowCheck {
    int getOffset();
    void setOffset(int offset);
    boolean isCheckSuccessful(HSSFRow row);
    boolean isCheckSuccessful(XLSRowCursor cursor);
    void addCellCheck(OffsetCellCheck cellCheck);
}
