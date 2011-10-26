package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.controller.SheetTransformationController;
import net.sf.jxls.controller.SheetTransformationControllerImpl;
import net.sf.jxls.tag.Block;
import net.sf.jxls.transformation.DuplicateTransformation;
import net.sf.jxls.transformation.RemoveTransformation;
import net.sf.jxls.transformation.ShiftTransformation;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.transformer.Workbook;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class SheetTransformationControllerTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());
    public static final String simpleBeanXLS = "/templates/simplebean.xls";
    public static final String duplicateOutputXLS = "target/duplicate_output.xls";
    public static final String removeBordersOutputXLS = "target/removeBorders_output.xls";
    public static final String removeBodyRowsXLS = "target/removeBodyRows_output.xls";

    public void testDuplicateDown() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        InputStream is1 = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));

        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(is);
        org.apache.poi.ss.usermodel.Sheet srcSheet = workbook.getSheetAt( 0 );
        org.apache.poi.ss.usermodel.Workbook destWorkbook = WorkbookFactory.create(is1);
        org.apache.poi.ss.usermodel.Sheet destSheet = destWorkbook.getSheetAt( 0 );
        int lastRowNum = srcSheet.getLastRowNum();
        Block block = new Block(null, 1, 3);
        Workbook wb = new Workbook(destWorkbook);
        Sheet sheet = new Sheet(destWorkbook, destSheet);
        wb.addSheet( sheet );
        wb.createFormulaController();
        SheetTransformationControllerImpl stc = new SheetTransformationControllerImpl(sheet);
        stc.duplicateDown( block, 2 );

//        assertEquals("Last Row Number is incorrect", lastRowNum + 3 * 2, destSheet.getLastRowNum());

        CellsChecker checker = new CellsChecker(new HashMap());
        checker.checkRows(srcSheet, destSheet, 0, 0, 4, true);
        checker.checkRows(srcSheet, destSheet, 1, 4, 3, true);
        checker.checkRows(srcSheet, destSheet, 1, 7, 3, true);
//        checker.checkRows(srcSheet, srcSheet, 4, 10, 1);
        // checking transformations
        List transformations = stc.getTransformations();
        assertEquals( "Number of transformations is incorrect", 2, transformations.size() );
        List expectedTransformations = new ArrayList();
        expectedTransformations.add( new ShiftTransformation( new Block(sheet, 4, Integer.MAX_VALUE), 6, 0));
        expectedTransformations.add( new DuplicateTransformation( new Block(sheet, 1, 3), 2));

        for( int i = 0; i < 2; i++ ){
            Object bt = transformations.get( i );
            Object ebt = expectedTransformations.get( i );
            assertEquals( "Transformation is incorrect", ebt, bt);
        }

        OutputStream os = new BufferedOutputStream(new FileOutputStream(duplicateOutputXLS));
        destWorkbook.write(os);
        is.close();
        os.flush();
        os.close();
    }

    public void testRemoveBorders() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        InputStream is1 = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));

        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(is);
        org.apache.poi.ss.usermodel.Sheet srcSheet = workbook.getSheetAt( 0 );
        org.apache.poi.ss.usermodel.Workbook destWorkbook = WorkbookFactory.create(is1);
        org.apache.poi.ss.usermodel.Sheet destSheet = destWorkbook.getSheetAt( 0 );
        int lastRowNum = srcSheet.getLastRowNum();
        Block block = new Block(null, 1, 3);
        Workbook wb = new Workbook(destWorkbook);
        Sheet sheet = new Sheet(destWorkbook, destSheet);
        wb.addSheet( sheet );
        wb.createFormulaController();

        SheetTransformationController stc = new SheetTransformationControllerImpl(sheet);
        stc.removeBorders( block );

        assertEquals("Last Row Number is incorrect", lastRowNum - 2, destSheet.getLastRowNum());

        CellsChecker checker = new CellsChecker(new HashMap());
        checker.checkRows(srcSheet, destSheet, 0, 0, 1, true);
        checker.checkRows(srcSheet, destSheet, 2, 1, 1, true);
        checker.checkRows(srcSheet, destSheet, 4, 2, 1, true);
        // checking transformations
        List transformations = stc.getTransformations();
        assertEquals( "Number of transformations is incorrect", 4, transformations.size() );
        List expectedTransformations = new ArrayList();
        expectedTransformations.add( new RemoveTransformation( new Block(sheet, 1, 1) ));
        expectedTransformations.add( new ShiftTransformation( new Block(sheet, 2, Integer.MAX_VALUE), -1, 0) );
        expectedTransformations.add( new RemoveTransformation( new Block(sheet, 2, 2) ));
        expectedTransformations.add( new ShiftTransformation( new Block(sheet, 3, Integer.MAX_VALUE), -1, 0) );

        expectedTransformations.add( new DuplicateTransformation( new Block(sheet, 1, 3), 2));

        for( int i = 0; i < 2; i++ ){
            Object bt = transformations.get( i );
            Object ebt = expectedTransformations.get( i );
            assertEquals( "Transformation is incorrect", ebt, bt);
        }

        OutputStream os = new BufferedOutputStream(new FileOutputStream(removeBordersOutputXLS));
        destWorkbook.write(os);
        is.close();
        os.flush();
        os.close();
    }

    public void testRemoveBodyRows() throws IOException, InvalidFormatException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        InputStream is1 = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));

        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(is);
        org.apache.poi.ss.usermodel.Sheet srcSheet = workbook.getSheetAt( 0 );
        org.apache.poi.ss.usermodel.Workbook destWorkbook = WorkbookFactory.create(is1);
        org.apache.poi.ss.usermodel.Sheet destSheet = destWorkbook.getSheetAt( 0 );
        int lastRowNum = srcSheet.getLastRowNum();
        Workbook wb = new Workbook(destWorkbook);
        Sheet sheet = new Sheet(destWorkbook, destSheet);
        Block block = new Block(sheet, 1, 3);
        wb.addSheet( sheet );
        wb.createFormulaController();
        SheetTransformationController stc = new SheetTransformationControllerImpl(sheet);
        stc.removeBodyRows( block );

        assertEquals("Last Row Number is incorrect", lastRowNum - 3, destSheet.getLastRowNum());

        CellsChecker checker = new CellsChecker(new HashMap());
        checker.checkRows(srcSheet, destSheet, 0, 0, 1, true);
//        checker.checkRows(srcSheet, destSheet, 2, 1, 1);
        checker.checkRows(srcSheet, destSheet, 4, 1, 1, true);
//        checker.checkRows(srcSheet, srcSheet, 4, 10, 1);
        // checking transformations
        List transformations = stc.getTransformations();
        assertEquals( "Number of transformations is incorrect", 2, transformations.size() );
        List expectedTransformations = new ArrayList();
        expectedTransformations.add( new RemoveTransformation( new Block(sheet, 1, 3) ));
        expectedTransformations.add( new ShiftTransformation( new Block(sheet, 4, Integer.MAX_VALUE), -3, 0) );

        expectedTransformations.add( new DuplicateTransformation( new Block(sheet, 1, 3), 2));

        for( int i = 0; i < 2; i++ ){
            Object bt = transformations.get( i );
            Object ebt = expectedTransformations.get( i );
            assertEquals( "Transformation is incorrect", ebt, bt);
        }

        OutputStream os = new BufferedOutputStream(new FileOutputStream(removeBodyRowsXLS));
        destWorkbook.write(os);
        is.close();
        os.flush();
        os.close();
    }
}
