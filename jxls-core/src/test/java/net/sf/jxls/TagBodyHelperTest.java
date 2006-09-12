package net.sf.jxls;

import junit.framework.TestCase;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import net.sf.jxls.util.TagBodyHelper;
import net.sf.jxls.tag.Block;
import net.sf.jxls.CellsChecker;

import java.util.Map;
import java.util.HashMap;
import java.io.*;

/**
 * @author Leonid Vysochyn
 */
public class TagBodyHelperTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());
    public static final String simpleBeanXLS = "/templates/simplebean.xls";
    public static final String simpeBeanDestXLS = "target/duplicate_output.xls";

    public static final String grouping1XLS = "/templates/grouping1.xls";
    public static final String grouping1DestXLS = "target/replace_output.xls";



    public void testDuplicateDown() throws IOException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        TagBodyHelper helper = new TagBodyHelper();

        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook workbook = new HSSFWorkbook(fs);
        HSSFSheet sheet = workbook.getSheetAt( 0 );
        int lastRowNum = sheet.getLastRowNum();
        Block block = new Block(null, 1, 3);
        helper.duplicateDown( sheet, block, 2);

        assertEquals("Last Row Number is incorrect", lastRowNum + 3 * 2, sheet.getLastRowNum());

        CellsChecker checker = new CellsChecker(new HashMap());
        checker.checkRows(sheet, sheet, 0, 0, 4);
        checker.checkRows(sheet, sheet, 1, 4, 3);
        checker.checkRows(sheet, sheet, 1, 7, 3);
//        checker.checkRows(sheet, sheet, 4, 10, 1);

        OutputStream os = new BufferedOutputStream(new FileOutputStream(simpeBeanDestXLS));
        workbook.write(os);
        is.close();
        os.flush();
        os.close();
    }

    public void testReplaceProperty() throws IOException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping1XLS));
        TagBodyHelper helper = new TagBodyHelper();

        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook workbook = new HSSFWorkbook(fs);
        HSSFSheet sheet = workbook.getSheetAt( 0 );
        int lastRowNum = sheet.getLastRowNum();
        Block block = new Block(null, 0, 4);
        helper.replaceProperty( sheet, block, "mainBean.beans", "item");

        assertEquals("Last Row Number is incorrect", lastRowNum, sheet.getLastRowNum());

        Map props = new HashMap();
        props.put( "mainBean.beans", "item");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sheet, sheet, 0, 0, 5);

        OutputStream os = new BufferedOutputStream(new FileOutputStream(grouping1DestXLS));
        workbook.write(os);
        is.close();
        os.flush();
        os.close();

    }

}
