package net.sf.jxls.reader;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import junit.framework.TestCase;
import net.sf.jxls.reader.sample.SimpleBean;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;

/**
 * @author Leonid Vysochyn
 * @version 1.0 25.07.2007
 */
public class ErrorHandlingTest extends TestCase {
    public static final String data1XLS = "/templates/error1.xls";
    public static final String xmlConfig1 = "/xml/error1.xml";

    public void testExceptionCatch() throws IOException, SAXException, InvalidFormatException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(data1XLS));
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig1));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        ReaderConfig.getInstance().setSkipErrors( false );
        Map beans = new HashMap();
        SimpleBean bean = new SimpleBean();
        beans.put( "bean", bean);
        try {
            reader.read( inputXLS,  beans);
            fail("Exception should be thrown");
        } catch (XLSDataReadException e) {
            System.out.println("Caught XLSDataReadException");
            assertNotNull( e.getReadStatus() );
            assertEquals("Number of ReadMessages is incorrect", 1, e.getReadStatus().getReadMessages().size());
            assertTrue("ReadStatus is incorrect", !e.getReadStatus().isStatusOK());
        }
    }

    public void testSkipErrors() throws IOException, SAXException, ParseException, InvalidFormatException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(data1XLS));
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig1));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        ReaderConfig.getInstance().setSkipErrors( true );
        Map beans = new HashMap();
        SimpleBean bean = new SimpleBean();
        beans.put( "bean", bean);

        XLSReadStatus readStatus = reader.read( inputXLS,  beans);
        assertEquals("Integer value read error", new Integer(5), bean.getIntValue3());
        SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy");
        Date date = format.parse("3/14/2007");
        assertEquals("Date value read error", date, bean.getDateValue());
        assertNotNull(readStatus);
        assertTrue( readStatus.isStatusOK() );
        assertEquals( "Number of ReadMessage object is incorrect", 3, readStatus.getReadMessages().size());
    }

}
