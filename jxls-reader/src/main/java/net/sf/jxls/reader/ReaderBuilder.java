package net.sf.jxls.reader;

import org.apache.commons.digester.Digester;
import org.xml.sax.SAXException;

import java.io.*;

/**
 * @author Leonid Vysochyn
 */
public class ReaderBuilder {

    XLSReader buildFromXML(InputStream xmlStream) throws IOException, SAXException {
        Digester digester = new Digester();
        digester.setValidating( false );
        digester.addObjectCreate("workbook", "net.sf.jxls.reader.XLSReaderImpl");
        digester.addObjectCreate("workbook/worksheet", "net.sf.jxls.reader.XLSSheetReaderImpl");
        digester.addSetProperties("workbook/worksheet", "name", "sheetName");
//        digester.addSetProperty("workbook/worksheet", "sheetName", "name");
        digester.addSetNext("workbook/worksheet", "addSheetReader");
        digester.addObjectCreate("*/loop", "net.sf.jxls.reader.XLSForEachBlockReaderImpl");
        digester.addSetProperties("*/loop");
        digester.addSetNext("*/loop", "addBlockReader");
        digester.addObjectCreate("*/section", "net.sf.jxls.reader.XLSBlockReaderImpl");
        digester.addSetProperties("*/section");
        digester.addSetNext("*/section", "addBlockReader");
        digester.addObjectCreate("*/mapping", "net.sf.jxls.reader.BeanCellMapping");
        digester.addSetProperties("*/mapping");
        digester.addCallMethod("*/mapping", "setFullPropertyName", 1);
        digester.addCallParam("*/mapping", 0);
        digester.addSetNext("*/mapping", "addMapping");
        digester.addObjectCreate("*/loop/loopbreakcondition", "net.sf.jxls.reader.SimpleSectionCheck");
        digester.addSetNext("*/loop/loopbreakcondition", "setLoopBreakCondition");
        digester.addObjectCreate("*/loopbreakcondition/rowcheck", "net.sf.jxls.reader.OffsetRowCheckImpl");
        digester.addSetProperties("*/loopbreakcondition/rowcheck");
        digester.addSetNext("*/loopbreakcondition/rowcheck", "addRowCheck");
        digester.addObjectCreate("*/rowcheck/cellcheck", "net.sf.jxls.reader.OffsetCellCheckImpl");
        digester.addSetProperties("*/rowcheck/cellcheck");
        digester.addCallMethod("*/rowcheck/cellcheck", "setValue", 1);
        digester.addCallParam("*/rowcheck/cellcheck", 0);
        digester.addSetNext("*/rowcheck/cellcheck", "addCellCheck");
        return (XLSReader) digester.parse( xmlStream );
    }

    XLSReader buildFromXML(File xmlFile) throws IOException, SAXException {
        InputStream xmlStream = new BufferedInputStream( new FileInputStream(xmlFile) );
        XLSReader reader = buildFromXML( xmlStream );
        xmlStream.close();
        return reader;
    }
}
