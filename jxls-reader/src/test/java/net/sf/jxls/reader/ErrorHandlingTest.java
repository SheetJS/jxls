package net.sf.jxls.reader;

import junit.framework.TestCase;
import net.sf.jxls.reader.sample.SimpleBean;
import org.xml.sax.SAXException;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 * @version 1.0 25.07.2007
 */
public class ErrorHandlingTest extends TestCase {
    public static final String data1XLS = "/templates/error1.xls";
    public static final String xmlConfig1 = "/xml/error1.xml";

    public void testExceptionCatch() throws IOException, SAXException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(data1XLS));
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig1));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        Map beans = new HashMap();
        SimpleBean bean = new SimpleBean();
        Configuration.getInstance().setUseDefaultValuesForPrimitiveTypes( false );
        beans.put( "bean", bean);
        try {
            reader.read( inputXLS,  beans);
            fail("Exception should be thrown");
        } catch (XLSDataReadException e) {
            System.out.println("Caught XLSDataReadException");
            e.printStackTrace();
        }
    }

    public void testSkipErrors() throws IOException, SAXException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(data1XLS));
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig1));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        Configuration.getInstance().setSkipErrors( true );
        Map beans = new HashMap();
        SimpleBean bean = new SimpleBean();
        beans.put( "bean", bean);
        try {
            reader.read( inputXLS,  beans);
            assertEquals("Integer value read error", new Integer(5), bean.getIntValue3());
        } catch (XLSDataReadException e) {
            System.out.println("Caught XLSDataReadException");
            e.printStackTrace();
        }
    }

}
