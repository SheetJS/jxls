package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.tag.Block;
import net.sf.jxls.util.TagBodyHelper;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class TagBodyHelperTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());
    public static final String simpleBeanXLS = "/templates/simplebean.xls";
    public static final String simpeBeanDestXLS = "target/duplicate_output.xls";

    public static final String grouping1XLS = "/templates/grouping1.xls";
    public static final String grouping1DestXLS = "target/replace_output.xls";



    public void testDuplicateDown() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        Workbook workbook = WorkbookFactory.create(is);
        Sheet sheet = workbook.getSheetAt( 0 );
        int lastRowNum = sheet.getLastRowNum();
        Block block = new Block(null, 1, 3);
        TagBodyHelper.duplicateDown( sheet, block, 2);

//        assertEquals("Last Row Number is incorrect", lastRowNum + 3 * 2, sheet.getLastRowNum());

        CellsChecker checker = new CellsChecker(new HashMap());
        checker.checkRows(sheet, sheet, 0, 0, 4, true);
        checker.checkRows(sheet, sheet, 1, 4, 3, true);
        checker.checkRows(sheet, sheet, 1, 7, 3, true);
//        checker.checkRows(sheet, sheet, 4, 10, 1);

        is.close();
        saveWorkbook( workbook, simpeBeanDestXLS);
    }

    public void testReplaceProperty() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping1XLS));
        Workbook workbook = WorkbookFactory.create(is);
        Sheet sheet = workbook.getSheetAt( 0 );
        int lastRowNum = sheet.getLastRowNum();
        Block block = new Block(null, 0, 4);
        TagBodyHelper.replaceProperty( sheet, block, "mainBean.beans", "item");

        assertEquals("Last Row Number is incorrect", lastRowNum, sheet.getLastRowNum());

        Map props = new HashMap();
        props.put( "mainBean.beans", "item");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sheet, sheet, 0, 0, 5, true);
        is.close();
        saveWorkbook( workbook, grouping1DestXLS);
    }

    private void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
        String saveResultsProp = System.getProperty("saveResults");
        if( "true".equalsIgnoreCase(saveResultsProp) ){
            OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
            resultWorkbook.write(os);
            os.flush();
            os.close();
        }
    }



}
