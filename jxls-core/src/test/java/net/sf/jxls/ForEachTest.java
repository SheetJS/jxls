package net.sf.jxls;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.XLSTransformer;
import net.sf.jxls.bean.Department;
import net.sf.jxls.bean.Employee;

import java.io.IOException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import junit.framework.TestCase;

/**
 * @author Leonid Vysochyn
 *         Date: 12.03.2009
 */
public class ForEachTest extends TestCase {
    public static final String forifTag2XLS = "/templates/foriftag2.xls";
    public static final String forifTag2DestXLS = "target/foriftag2_output.xls";

    public static final String forifTag3XLS = "/templates/foriftag3.xls";
    public static final String forifTag3DestXLS = "target/foriftag3_output.xls";
    

    List itEmployees = new ArrayList();

    String[] itEmployeeNames = new String[] {"Elsa", "Oleg", "Neil", "Maria", "John"};
    String[] hrEmployeeNames = new String[] {"Olga", "Helen", "Keith", "Cat"};
    String[] baEmployeeNames = new String[] {"Denise", "LeAnn", "Natali"};
    String[] mgrEmployeeNames = new String[] {"Sean", "John", "Joerg"};

    Double[] itPayments = new Double[] {new Double(1500), new Double(2300), new Double(2500), new Double(1700), new Double(2800)};
    Double[] hrPayments = new Double[] {new Double(1400), new Double(2100), new Double(1800), new Double(1900)};
    Double[] baPayments = new Double[] {new Double(2400), new Double(2200), new Double(2600)};
    Double[] mgrPayments = new Double[] {null, new Double(6000), null};
    Double[] itBonuses = new Double[] {new Double(0.15), new Double(0.25), new Double(0.00), new Double(0.15), new Double(0.20)};
    Double[] hrBonuses = new Double[] {new Double(0.20), new Double(0.10), new Double(0.15), new Double(0.15)};
    Double[] baBonuses = new Double[] {new Double(0.20), new Double(0.15), new Double(0.10)};
    Double[] mgrBonuses = new Double[] {new Double(0.20), null, new Double(0.20)};
    Integer[] itAges = new Integer[] {new Integer(34), new Integer(30), new Integer(25), new Integer(25), new Integer(35)};
    Integer[] hrAges = new Integer[] {new Integer(26), new Integer(28), new Integer(26), new Integer(26)};
    Integer[] baAges = new Integer[] {new Integer(30), new Integer(30), new Integer(30)};
    Integer[] mgrAges = new Integer[] {null, new Integer(35), null};
    List departments = new ArrayList();
    Department mgrDepartment, itDepartment;

    protected void setUp() throws Exception {
        super.setUp();
        Department department = new Department("IT");
        for(int i = 0; i < itEmployeeNames.length; i++){
            Employee employee = new Employee(itEmployeeNames[i], itAges[i], itPayments[i], itBonuses[i]);
            employee.setNotes( generateNotes(employee.getName()) );
            department.addEmployee( employee );
            itEmployees.add( employee );
        }
        itDepartment = department;
        departments.add( department );
        department = new Department("HR");
        for(int i = 0; i < hrEmployeeNames.length; i++){
            department.addEmployee( new Employee(hrEmployeeNames[i], hrAges[i], hrPayments[i], hrBonuses[i]) );
        }
        departments.add( department );
        department = new Department("BA");
        for(int i = 0; i < baEmployeeNames.length; i++){
            department.addEmployee( new Employee(baEmployeeNames[i], baAges[i], baPayments[i], baBonuses[i]) );
        }
        departments.add( department );

        department = new Department("MGR");
        for(int i = 0; i < mgrEmployeeNames.length; i++){
            department.addEmployee( new Employee(mgrEmployeeNames[i], mgrAges[i], mgrPayments[i], mgrBonuses[i]) );
        }
        mgrDepartment = department;
    }

    protected List generateNotes(String name) {
        Random r = new Random( System.currentTimeMillis() );
        int n = 1 + r.nextInt(7);
        List notes = new ArrayList();
        for(int i = 0 ; i < n; i++){
            notes.add("Note " + i + " for " + name);
        }
        return notes;
    }


    public void testForIfTag2() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );
        beans.put("depUrl", "http://www.somesite.com");

        Configuration config = new Configuration();
        config.setMetaInfoToken("\\\\");

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTag2XLS));
        XLSTransformer transformer = new XLSTransformer( config );
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag2XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", 11, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${department.name}", "IT");
        props.put("${depUrl}", "http://www.somesite.com");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 0, itEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 1, itPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 2, itBonuses);
//todo:        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 3, (short)3, "B4*(1+C4)");
//todo:        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 4, (short)3, "B5*(1+C5)");
//todo:        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 7, (short)3, "B8*(1+C8)");
      //        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 8, (short)1, "SUM(B4:B8)");
          //        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 6, (short)3, "SUM(D4:D8)");
        props.clear();
        props.put("${department.name}", "HR");
        props.put("${depUrl}", "http://www.somesite.com");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 9, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 0, hrEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 1, hrPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 2, hrBonuses);
//todo        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 12, (short)3, "B13*(1+C13)");
//todo        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 13, (short)3, "B14*(1+C14)");
//todo        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 15, (short)3, "B16*(1+C16)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 16, (short)1, "SUM(B13:B16)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 16, (short)3, "SUM(D13:D16)");
        props.clear();
        props.put("${department.name}", "BA");
        props.put("${depUrl}", "http://www.somesite.com");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 17, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 0, baEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 1, baPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 2, baBonuses);
        //todo:
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 20, (short)3, "B21*(1+C21)");
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 21, (short)3, "B22*(1+C22)");
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 22, (short)3, "B23*(1+C23)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 23, (short)1, "SUM(B21:B23)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 23, (short)3, "SUM(D21:D23)");

        is.close();
    }

    public void testForIfTag3() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );
        beans.put("depUrl", "http://www.somesite.com");
        List deps = new ArrayList();
        Department testDep = new Department("Test");
        deps.add( testDep );
        beans.put( "deps", deps );
        List employees = new ArrayList();
        beans.put("employees", employees);

        Configuration config = new Configuration();
        config.setMetaInfoToken("\\\\");

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3XLS));
        XLSTransformer transformer = new XLSTransformer( config );
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", 54, resultSheet.getLastRowNum());

        // check 1st forEach loop output
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 4, 3);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 8, 3);
        checker.checkRows(sourceSheet, resultSheet, 11, 3, 1);
        checker.checkRows(sourceSheet, resultSheet, 11, 7, 1);
        checker.checkRows(sourceSheet, resultSheet, 11, 11, 1);
        // check 2nd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 12, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 11, 18, 1);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 19, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 23, 1);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 24, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 30, 1);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 14, 12, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 22, 18, 1);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 14, 19, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 23, 1);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 14, 24, 3);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 30, 1);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 25, 31, 3);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)0, itEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)1, itPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)2, itBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 18, 1);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 25, 40, 3);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)0, hrEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)1, hrPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)2, hrBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 23, 1);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 25, 48, 3);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)0, baEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)1, baPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)2, baBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 30, 1);
        sourceSheet = sourceWorkbook.getSheetAt( 1 );
        resultSheet = resultWorkbook.getSheetAt( 1 );
        assertEquals("Number of rows on Sheet 2 is not correct", 1, resultSheet.getLastRowNum() + 1);
        checker.setIgnoreFirstLastCellNums( true );
        checker.checkRows( sourceSheet, resultSheet, 11, 0, 1);
        is.close();
    }

}

