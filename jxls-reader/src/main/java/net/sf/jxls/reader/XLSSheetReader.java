package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.List;
import java.util.Map;

/**
 * Interface to parse XLS sheet 
 * @author Leonid Vysochyn
 */
public interface XLSSheetReader {
    /**
     * Method to read data from excel sheet and populate objects
     * @param sheet - {@link HSSFSheet} object
     * @param beans - {@link Map} of beans to populate
     * @return {@link XLSReadStatus} object with info about read status
     */
    XLSReadStatus read(HSSFSheet sheet, Map beans);

    List getBlockReaders();
    void setBlockReaders(List blockReaders);
    void addBlockReader(XLSBlockReader blockReader);

    String getSheetName();
    void setSheetName(String sheetName);

}
