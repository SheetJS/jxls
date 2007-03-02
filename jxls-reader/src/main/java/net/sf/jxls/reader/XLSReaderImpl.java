package net.sf.jxls.reader;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Basic implementation of {@link XLSReader} interface
 * @author Leonid Vysochyn
 */
public class XLSReaderImpl implements XLSReader {
    protected final Log log = LogFactory.getLog(getClass());

    Map sheetReaders = new HashMap();

    public void read(InputStream inputXLS, Map beans) throws IOException {
        POIFSFileSystem fsInput = new POIFSFileSystem(inputXLS);
        HSSFWorkbook workbook = new HSSFWorkbook(fsInput);
        for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
            HSSFSheet sheet = workbook.getSheetAt( sheetNo );
            String sheetName = workbook.getSheetName( sheetNo );
            if( log.isInfoEnabled() ){
                log.info("Processing sheet " + sheetName);
            }
            if( sheetReaders.containsKey( sheetName ) ){
                XLSSheetReader sheetReader = (XLSSheetReader) sheetReaders.get( sheetName );
                sheetReader.setSheetName( sheetName );
                sheetReader.read( sheet, beans );
            }
        }
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
