package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class XLSSheetReaderImpl implements XLSSheetReader {

    List blockReaders = new ArrayList();
    String sheetName;

    public void read(HSSFSheet sheet, Map beans) {
        XLSRowCursor cursor = new XLSRowCursorImpl( sheetName, sheet );
        for (int i = 0; i < blockReaders.size(); i++) {
            XLSBlockReader blockReader = (XLSBlockReader) blockReaders.get(i);
            blockReader.read( cursor, beans );
            cursor.moveForward();
        }
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
