package net.sf.jxls.reader;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Leonid Vysochyn
 */
public class XLSSheetReaderImpl implements XLSSheetReader {

    List blockReaders = new ArrayList();
    String sheetName;

    XLSReadStatus readStatus = new XLSReadStatus();


    public XLSReadStatus read(Sheet sheet, Map beans) {
        readStatus.clear();
        XLSRowCursor cursor = new XLSRowCursorImpl( sheetName, sheet );
        for (int i = 0; i < blockReaders.size(); i++) {
            XLSBlockReader blockReader = (XLSBlockReader) blockReaders.get(i);
            readStatus.mergeReadStatus( blockReader.read( cursor, beans ) );
            cursor.moveForward();
        }
        return readStatus;
    }

    public List getBlockReaders() {
        return blockReaders;
    }

    public void setBlockReaders(List blockReaders) {
        this.blockReaders = blockReaders;
    }

    public void addBlockReader(XLSBlockReader blockReader) {
        blockReaders.add( blockReader );
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}
