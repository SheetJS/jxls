package net.sf.jxls.reader;

import java.util.List;
import java.util.Map;

/**
 * Interface to read block of excel rows
 * @author Leonid Vysochyn
 */
public interface XLSBlockReader {
    void read(XLSRowCursor cursor, Map beans);
    void setLoopBreakCondition(SectionCheck condition);
    SectionCheck getLoopBreakCondition();
    void addBlockReader(XLSBlockReader reader);
    List getBlockReaders();
    int getStartRow();
    void setStartRow(int startRow);
    int getEndRow();
    void setEndRow(int endRow);
    public void addMapping(BeanCellMapping mapping);
}
