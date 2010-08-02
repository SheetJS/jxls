package net.sf.jxls.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Basic implementation of {@link XLSReader} interface
 * @author Leonid Vysochyn
 */
public class XLSReaderImpl implements XLSReader {
    protected final Log log = LogFactory.getLog(getClass());

    Map sheetReaders = new HashMap();

    XLSReadStatus readStatus = new XLSReadStatus();


    public XLSReadStatus read(InputStream inputXLS, Map beans) throws IOException, InvalidFormatException {
        readStatus.clear();
        Workbook workbook = WorkbookFactory.create(inputXLS);
        for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
            readStatus.mergeReadStatus( readSheet(workbook, sheetNo, beans) );
        }
        return readStatus;
    }

    private XLSReadStatus readSheet(Workbook workbook, int sheetNo, Map beans) {
        Sheet sheet = workbook.getSheetAt( sheetNo );
        String sheetName = workbook.getSheetName( sheetNo );
        if( log.isInfoEnabled() ){
            log.info("Processing sheet " + sheetName);
        }
        if( sheetReaders.containsKey( sheetName ) ){
            XLSSheetReader sheetReader = (XLSSheetReader) sheetReaders.get( sheetName );
            sheetReader.setSheetName( sheetName );
            return sheetReader.read( sheet, beans );
        }
        return null;
    }

    public Map getSheetReaders() {
        return sheetReaders;
    }

    public void addSheetReader(String sheetName, XLSSheetReader reader) {
        sheetReaders.put( sheetName, reader );
    }

    public void addSheetReader(XLSSheetReader reader) {
        addSheetReader( reader.getSheetName(), reader );
    }

    public void setSheetReaders(Map sheetReaders) {
        this.sheetReaders = sheetReaders;
    }
}
