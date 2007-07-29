package net.sf.jxls.reader;

import java.util.Map;

/**
 * Interface to read block of excel rows
 * @author Leonid Vysochyn
 */
public interface XLSBlockReader {
    XLSReadStatus read(XLSRowCursor cursor, Map beans);

    int getStartRow();

    void setStartRow(int startRow);

    int getEndRow();

    void setEndRow(int endRow);

}
