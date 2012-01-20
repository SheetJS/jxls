package net.sf.jxls;

import net.sf.jxls.formula.CellRef;
import net.sf.jxls.formula.Formula;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Test case for Formula class
 */
public class FormulaTest extends BaseTest {
    public static final String formulaXLS = "/templates/formula3.xls";
    public static final String formulaDestXLS = "target/formula3_output.xls";

    public static final String formula4XLS = "/templates/formula4.xls";
    public static final String formula4DestXLS = "target/formula4_output.xls";

    public static final String formulaOneRowXLS = "/templates/formulaOneRow.xlsx";
    public static final String formulaOneRowDestXLS = "target/formulaOneRow_output.xlsx";

    public void testFormulaOneRowForEach() throws IOException, InvalidFormatException {
        Map values = new HashMap();
        values.put("list", new Double[]{10.5d, 20d, 30d, 40.5d});
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formulaOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, values);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkFormulaCell(resultSheet, 0, 4, "SUM(A1:D1)");
        is.close();
        saveWorkbook(resultWorkbook, formulaOneRowDestXLS);
    }


    public void testFindRefCells(){
        String formulaValue = "SUM(a1:a10) - D12 + C5 * D10 - 4 + MULT ( B2 : B90 )";
        Formula formula = new Formula( formulaValue, null );
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
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkCell(resultSheet, 8, 1, 1500);
        checker.checkCell(resultSheet, 9, 1, 2300);
        checker.checkFormulaCell(resultSheet, 13, 1, "SUM(B9:B13)");
        is.close();
        saveWorkbook(resultWorkbook, formulaDestXLS);
    }

    public void testMultipleSheets() throws InvalidFormatException, IOException {
        Map beans = new HashMap();
        List d1List = new ArrayList();
        for(int i = 0; i <5; i++){
            Map d1 = new HashMap();
            d1.put("clt_category", "cat " + i);
            d1.put("cac_desc", "desc " + i);
            d1.put("numnew", i);
            d1.put("totnew", i*2);
            d1.put("tienew", i*2);
            d1.put("tietot", i*3);
            d1List.add(d1);
        }
        List d2List = new ArrayList();
        for(int i=0; i<3; i++){
            Map d2 = new HashMap();
            d2.put("clt_no", "no " + i);
            d2.put("clt_name_1", "name " + i);
            d2.put("numnew", i);
            d2.put("totnew", i * 2);
            d2.put("tienew", i * 2);
            d2.put("tietot", i * 3);
            d2List.add(d2);
        }

        List d3List = new ArrayList();
        for(int i=0; i<7; i++){
            Map d3 = new HashMap();
            d3.put("clt_category", "cat " + i);
            d3.put("cac_desc", "desc " + i);
            d3.put("addnew", i);
            d3.put("addtot", i * 2);
            d3.put("new1tr", i * 2);
            d3.put("tot1tr", i * 3);
            d3List.add(d3);
        }

        List d4List = new ArrayList();
        for(int i=0; i<3; i++){
            Map d4 = new HashMap();
            d4.put("clt_no", "no " + i);
            d4.put("clt_name_1", "name " + i);
            d4.put("addnew", i);
            d4.put("addtot", i * 2);
            d4.put("new1tr", i * 2);
            d4.put("tot1tr", i * 3);
            d4List.add(d4);
        }

        List d5List = new ArrayList();
        for(int i=0; i<3; i++){
            Map d5 = new HashMap();
            d5.put("a", "name " + i);
            d5.put("b", i);
            d5.put("c", i*2);
            d5.put("d", i * 3);
            d5List.add(d5);
        }

        List d6List = new ArrayList();
        for(int i=0; i<6; i++){
            Map d6 = new HashMap();
            d6.put("a", "r " + i);
            d6.put("b", i);
            d6.put("c", i * 2);
            d6.put("d", i * 3);
            d6List.add(d6);
        }

        List d7List = new ArrayList();
        for(int i=0; i<2; i++){
            Map d7 = new HashMap();
            d7.put("a", "s " + i);
            d7.put("b", i);
            d7.put("c", i * 4);
            d7.put("d", i * 5);
            d7List.add(d7);
        }
        List d8List = new ArrayList();
        for(int i=0; i<5; i++){
            Map d8 = new HashMap();
            d8.put("a", "t " + i);
            d8.put("b", i);
            d8.put("c", i * 7);
            d8.put("d", i * 9);
            d8List.add(d8);
        }


        beans.put( "d1", d1List );
        beans.put( "d2", d2List );
        beans.put( "d3", d3List );
        beans.put( "d4", d4List );
        beans.put( "d5", d5List );
        beans.put( "d6", d6List );
        beans.put( "d7", d7List );
        beans.put( "d8", d8List );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formula4XLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        Sheet sheet0 = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkFormulaCell(sheet0, 10, 2, "SUM(C6:C10)");
        checker.checkFormulaCell(sheet0, 10, 3, "SUM(D6:D10)");
        checker.checkFormulaCell(sheet0, 17, 2, "SUM(C15:C17)");
        Sheet sheet1 = resultWorkbook.getSheetAt(1);
        checker.checkFormulaCell(sheet1, 12, 2, "SUM(C6:C12)");
        checker.checkFormulaCell(sheet1, 12, 3, "SUM(D6:D12)");
        checker.checkFormulaCell(sheet1, 19, 2, "SUM(C17:C19)");
        checker.checkFormulaCell(sheet1, 19, 3, "SUM(D17:D19)");
        is.close();
        saveWorkbook(resultWorkbook, formula4DestXLS);

    }

}
