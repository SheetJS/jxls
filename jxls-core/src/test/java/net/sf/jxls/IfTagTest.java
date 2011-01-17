package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.bean.SimpleBean;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * @author Leonid Vysochyn
 *         Date: 09.04.2010
 */
public class IfTagTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());

    public static final String ifTagEmptyXLS = "/templates/iftagempty.xls";
    public static final String ifTagEmptyDestXLS = "target/iftagempty_output.xls";

    public static final String twoIfTagsIn1RowXLS = "/templates/twoIfTagsIn1Row.xls";
    public static final String twoIfTagsIn1RowDestXLS = "target/twoIfTagsIn1Row_output.xls";


    public void testEmptyCollection() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        List items = new ArrayList();
        items.add(new SimpleBean("Simple bean"));
        beans.put( "items", items );
        beans.put("emptyItems", new ArrayList());
        beans.put("nullItems", null);

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(ifTagEmptyXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkRow(resultSheet, 0, 0, 1, new Object[]{"Name:", "Simple bean"});
        checker.checkRow(resultSheet, 2, 0, 0, new Object[]{"This collection is empty"});
        checker.checkRow(resultSheet, 4, 0, 0, new Object[]{"This collection detected as null"});
//        Object[] values = new Object[]{"IT", "IT", null, "Elsa", new Double(1500), "Oleg", new Double(2300),
//                "Neil", new Double(2500), "Maria", new Double(1700), "John", new Double(2800), "IT", "IT", "IT"};
//        checker.checkRow(resultSheet, 0, 0, 13, values);
        saveWorkbook(resultWorkbook, ifTagEmptyDestXLS);
    }

    private void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
        String saveResultsProp = System.getProperty("saveResults");
        if ("true".equalsIgnoreCase(saveResultsProp)) {
            if (log.isInfoEnabled()) {
                log.info("Saving " + fileName);
            }
            OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
            resultWorkbook.write(os);
            os.flush();
            os.close();
            log.info("Output Excel saved to " + fileName);
        }
    }

    public void test2IfTagsIn1Row() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(twoIfTagsIn1RowXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, new HashMap());
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkRow(resultSheet, 0, 0, 25, new Object[]{1, 2, 3, 5, 6,	7,	9,	10,	11,	12,	13,	14,	16,	17, 18,	19,	21,	22, 23,	24,	25,	26,	27,	28,	29,	30,	31});
        saveWorkbook(resultWorkbook, twoIfTagsIn1RowDestXLS);
    }

}
