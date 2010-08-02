package net.sf.jxls.reader;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author Leonid Vysochyn
 */
public interface OffsetRowCheck {
    int getOffset();
    void setOffset(int offset);
    boolean isCheckSuccessful(Row row);
    boolean isCheckSuccessful(XLSRowCursor cursor);
    void addCellCheck(OffsetCellCheck cellCheck);
}
