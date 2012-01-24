package net.sf.jxls.reader;

import junit.framework.TestCase;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class EmptyRowTest extends TestCase {
    public static final String dataXLS = "/templates/emptyrowdata.xls";
    public static final String xmlConfig = "/xml/emptyrow.xml";

    public void testReadList() throws IOException, SAXException, InvalidFormatException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(dataXLS));
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig));
        XLSReader reader = ReaderBuilder.buildFromXML(inputXML);
        ReaderConfig.getInstance().setSkipErrors( true );
        Map beans = new HashMap();
        List rules = new ArrayList();
        beans.put("rules", rules);
        try {
            reader.read(inputXLS, beans);
            inputXLS.close();
            assertNotNull(rules);
            assertEquals(1, rules.size());
        } catch (XLSDataReadException e) {
            e.printStackTrace();
            fail("No exception should be thrown");
        } 
    }
}
