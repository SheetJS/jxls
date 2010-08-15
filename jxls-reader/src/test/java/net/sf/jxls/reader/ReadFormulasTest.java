package net.sf.jxls.reader;

import junit.framework.TestCase;

import java.io.IOException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;
import net.sf.jxls.reader.sample.SimpleBean;

/**
 * @author Leonid Vysochyn
 *         Created: Feb 19, 2008
 */
public class ReadFormulasTest extends TestCase {
    public static final String formulasXLS = "/templates/formulasData.xls";
    public static final String xmlConfig = "/xml/formulas.xml";

    public void testReadFormulas() throws IOException, SAXException, InvalidFormatException {
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream( xmlConfig ));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        assertNotNull( reader );
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(formulasXLS));
        List employees = new ArrayList();
        SimpleBean bean = new SimpleBean();
        Map beans = new HashMap();
        beans.put("employees", employees);
        beans.put("bean", bean);
        reader.read( inputXLS, beans);
        assertNotNull( employees );
        assertEquals(4, employees.size());
        assertEquals( "Value or formula is incorrect ", new Integer(5), bean.getIntValue1() );
        assertEquals( "Value or formula is incorrect ", new Double(9805), bean.getDoubleValue() );
        assertEquals( "Value or formula is incorrect ", "Age&Payment", bean.getStr() );
        inputXLS.close();

    }

}
