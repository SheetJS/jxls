package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;

/**
 * @author Leonid Vysochyn
 */
public interface OffsetCellCheck {
    Object getValue();
    void setValue(Object value);
    short getOffset();
    void setOffset(short offset);
    boolean isCheckSuccessful(HSSFCell cell);
    boolean isCheckSuccessful(HSSFRow row);
}
