package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public interface SectionCheck {
    boolean isCheckSuccessful(XLSRowCursor cursor);
    void addRowCheck(OffsetRowCheck offsetRowCheck);
}
