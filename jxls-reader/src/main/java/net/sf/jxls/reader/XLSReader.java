package net.sf.jxls.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Interface to read and parse excel file
 * @author Leonid Vysochyn
 */
public interface XLSReader {
    XLSReadStatus read(InputStream inputXLS, Map beans) throws IOException;
    void setSheetReaders(Map sheetReaders);
    Map getSheetReaders();
    void addSheetReader( String sheetName, XLSSheetReader reader);
    void addSheetReader(XLSSheetReader reader);
}
