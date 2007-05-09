package net.sf.jxls.reader;

import java.util.List;

/**
 * Interface to read repetitive block of excel rows
 * @author Leonid Vysochyn
 */
public interface XLSLoopBlockReader extends XLSBlockReader {
    void setLoopBreakCondition(SectionCheck condition);
    SectionCheck getLoopBreakCondition();
    void addBlockReader(XLSBlockReader reader);
    List getBlockReaders();
}
