package net.sf.jxls;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;
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
public class XlsxTest extends BaseTest {

    public static final String simpleXlsx = "/templates/simple.xlsx";
    public static final String simpleDestXLSX = "target/simple_output.xlsx";


    public void testWriteXlsx() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleXlsx));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(simpleXlsx));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);

        Map props = new HashMap();
        props.put("${department.name}", "IT");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3, false);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 0, itEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 1, itPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 2, itBonuses);
        checker.ignoreStyle = true;
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 3, (short)3, "B4*(1+C4)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 5, (short)3, "B6*(1+C6)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 7, (short)3, "B8*(1+C8)");
        checker.checkFormulaCell( sourceSheet, 7, resultSheet, 8, (short)1, "SUM(B4:B8)");
        checker.ignoreStyle = false;
        props.clear();
        props.put("${department.name}", "HR");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 9, 3, false);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 0, hrEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 1, hrPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 2, hrBonuses);
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 12, (short)3, "B13*(1+C13)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 13, (short)3, "B14*(1+C14)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 15, (short)3, "B16*(1+C16)");
        checker.checkFormulaCell( sourceSheet, 7, resultSheet, 16, (short)1, "SUM(B13:B16)");
        props.clear();
        props.put("${department.name}", "BA");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 17, 3, false);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 0, baEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 1, baPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 2, baBonuses);
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 20, (short)3, "B21*(1+C21)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 21, (short)3, "B22*(1+C22)");
        checker.checkFormulaCell( sourceSheet, 5, resultSheet, 22, (short)3, "B23*(1+C23)");
        checker.checkFormulaCell( sourceSheet, 7, resultSheet, 23, (short)1, "SUM(B21:B23)");
        saveWorkbook(resultWorkbook, simpleDestXLSX);
    }

}
