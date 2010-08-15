package net.sf.jxls;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import junit.framework.TestCase;
import net.sf.jxls.formula.CellRef;
import net.sf.jxls.formula.Formula;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Test case for Formula class
 */
public class FormulaTest extends BaseTest {
    public static final String formulaXLS = "/templates/formula3.xls";
    public static final String formulaDestXLS = "target/formula3_output.xls";


    public void testFindRefCells(){
        String formulaValue = "SUM(a1:a10) - D12 + C5 * D10 - 4 + MULT ( B2 : B90 )";
        Formula formula = new Formula( formulaValue );
        Set refCells = formula.findRefCells();
        assertEquals( "Incorrect number of ref cells found", refCells.size(), 7  );
        assertTrue( contains(refCells, "a1") );
        assertTrue( contains(refCells, "a10" ) );
        assertTrue( contains(refCells, "D12" ) );
        assertTrue( contains(refCells, "C5" ) );
        assertTrue( contains(refCells, "D10" ) );
        assertTrue( contains(refCells, "B2" ) );
        assertTrue( contains(refCells, "B90" ) );
    }

    boolean contains(Set refCells, String cellRef){
        for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
            CellRef ref = (CellRef) iterator.next();
            if( ref.toString().equals( cellRef ) ){
                return true;
            }
        }
        return false;
    }

    public void testFormulaWhenTopRowsAreNull() throws InvalidFormatException, IOException {
        Map beans = new HashMap();
        beans.put( "department", itDepartment );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formulaXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
//        is = new BufferedInputStream(getClass().getResourceAsStream(formulaXLS));
//        Workbook sourceWorkbook = WorkbookFactory.create(is);
//
//        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
//        Sheet resultSheet = resultWorkbook.getSheetAt(0);
//
//        Map props = new HashMap();
//        props.put("${department.name}", "IT");
//        CellsChecker checker = new CellsChecker(props);
//        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3, false);
        saveWorkbook(resultWorkbook, formulaDestXLS);
    }

}
