package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.bean.Department;
import net.sf.jxls.bean.Employee;
import net.sf.jxls.bean.Feedback;
import net.sf.jxls.bean.FeedbackCount;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author Leonid Vysochyn
 *         Date: 12.03.2009
 */
public class ForEachTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());

    public static final String forifTag2XLS = "/templates/foriftag2.xls";
    public static final String forifTag2DestXLS = "target/foriftag2_output.xls";

    public static final String forifTag3XLS = "/templates/foriftag3.xls";
    public static final String forifTag3DestXLS = "target/foriftag3_output.xls";

    public static final String forifTag3OutTagXLS = "/templates/foriftag3OutTag.xls";
    public static final String forifTag3OutTagDestXLS = "target/foriftag3OutTag_output.xls";

    public static final String forifTagMergeXLS = "/templates/foriftagmerge.xls";
    public static final String forifTagMergeDestXLS = "target/foriftagmerge_output.xls";

    public static final String forifTagOneRowXLS = "/templates/foriftagOneRow.xls";
    public static final String forifTagOneRowDestXLS = "target/foriftagOneRow_output.xls";

    public static final String forOneRowXLS = "/templates/forOneRow.xls";
    public static final String forOneRowDestXLS = "target/forOneRow_output.xls";

    public static final String forOneRowMergeXLS = "/templates/forOneRowMerge.xls";
    public static final String forOneRowMergeDestXLS = "target/forOneRowMerge_output.xls";

    public static final String forOneRowMerge2XLS = "/templates/forOneRowMerge2.xls";
    public static final String forOneRowMerge2DestXLS = "target/forOneRowMerge2_output.xls";

    public static final String doubleForEachOneRowXLS = "/templates/doubleForEachOneRow.xls";
    public static final String doubleForEachOneRowDestXLS = "target/doubleForEachOneRow_output.xls";

    public static final String forGroupByXLS = "/templates/forgroup.xls";
    public static final String forGroupByDestXLS = "target/forgroup_output.xls";

    public static final String grouping4XLS = "/templates/grouping4.xls";
    public static final String grouping4DestXLS = "target/grouping4_output.xls";

    public static final String selectXLS = "/templates/select.xls";
    public static final String selectDestXLS = "/templates/select_output.xls";

    public static final String outTagOneRowXLS = "/templates/outtaginonerow.xls";
    public static final String outTagOneRowDestXLS = "/templates/outtaginonerow_output.xls";

    public static final String arrayXLS = "/templates/array.xls";
    public static final String arrayDestXLS = "target/array_output.xls";


    public static final String varStatusXLS = "/templates/varstatus.xls";
    public static final String varStatusDestXLS = "target/varstatus_output.xls";

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


    public void testVarStatusAttrInForEach() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put("employees", itEmployees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(varStatusXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(varStatusXLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        checker.checkListCells(sourceSheet, 3, resultSheet, 2, (short) 0, new Object[]{new Integer(0), new Integer(1), new Integer(2), new Integer(3), new Integer(4)});
        is.close();
        saveWorkbook(resultWorkbook, varStatusDestXLS);
    }


    public void testForIfTag2() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "departments", departments );
        beans.put("depUrl", "http://www.somesite.com");

        Configuration config = new Configuration();
        config.setMetaInfoToken("\\\\");

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTag2XLS));
        XLSTransformer transformer = new XLSTransformer( config );
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag2XLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", 11, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${department.name}", "IT");
        props.put("${depUrl}", "http://www.somesite.com");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3, true);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 0, itEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 1, itPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 2, itBonuses);
        props.clear();
        props.put("${department.name}", "HR");
        props.put("${depUrl}", "http://www.somesite.com");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 9, 3, true);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 0, hrEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 1, hrPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 2, hrBonuses);
        props.clear();
        props.put("${department.name}", "BA");
        props.put("${depUrl}", "http://www.somesite.com");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 17, 3, true);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 0, baEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 1, baPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 2, baBonuses);

        is.close();
    }

    public void testForIfTag3() throws IOException, ParsePropertyException, InvalidFormatException {
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
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3XLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());

        // check 1st forEach loop output
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 4, 3, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 8, 3, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 3, 1, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 7, 1, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 11, 1, true);
        // check 2nd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 12, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 11, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 19, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 24, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 30, 1, true);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 14, 12, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 22, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 14, 19, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 14, 24, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 30, 1, true);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 25, 31, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)0, itEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)1, itPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)2, itBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 25, 40, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)0, hrEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)1, hrPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)2, hrBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 25, 48, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)0, baEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)1, baPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)2, baBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 30, 1, true);
        sourceSheet = sourceWorkbook.getSheetAt( 1 );
        resultSheet = resultWorkbook.getSheetAt( 1 );
        assertEquals("Number of rows on Sheet 2 is not correct", 1, resultSheet.getLastRowNum() + 1);
        checker.setIgnoreFirstLastCellNums( true );
        checker.checkRows( sourceSheet, resultSheet, 11, 0, 1, true);
        is.close();
    }

    public void testForIfTag3OutTag() throws IOException, ParsePropertyException, InvalidFormatException {
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

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3OutTagXLS));
        XLSTransformer transformer = new XLSTransformer( config );
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3OutTagXLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);

        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());

        // check 1st forEach loop output
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 4, 3, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 8, 3, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 3, 1, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 7, 1, true);
        checker.checkRows(sourceSheet, resultSheet, 11, 11, 1, true);
        // check 2nd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 1, 12, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 11, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 1, 19, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 1, 24, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 11, 30, 1, true);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 14, 12, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)0, new String[]{"Oleg", "Neil", "John"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)1, new Double[]{new Double(2300), new Double(2500), new Double(2800)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 15, (short)2, new Double[]{new Double(0.25), new Double(0.00), new Double(0.20)});
        checker.checkRows(sourceSheet, resultSheet, 22, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 14, 19, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)0, new String[]{"Helen"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)1, new Double[]{new Double(2100)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 22, (short)2, new Double[]{new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 14, 24, 3, true);
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)0, new String[]{"Denise", "LeAnn", "Natali"});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)1, new Double[]{new Double(2400), new Double(2200), new Double(2600)});
        checker.checkListCells( sourceSheet, 19, resultSheet, 27, (short)2, new Double[]{new Double(0.20),new Double(0.15),new Double(0.10)});
        checker.checkRows(sourceSheet, resultSheet, 22, 30, 1, true);
        // check 3rd forEach loop output
        props.put("${department.name}", "IT");
        checker.checkRows(sourceSheet, resultSheet, 25, 31, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)0, itEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)1, itPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 34, (short)2, itBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 18, 1, true);
        props.put("${department.name}", "HR");
        checker.checkRows(sourceSheet, resultSheet, 25, 40, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)0, hrEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)1, hrPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 43, (short)2, hrBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 23, 1, true);
        props.put("${department.name}", "BA");
        checker.checkRows(sourceSheet, resultSheet, 25, 48, 3, true);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)0, baEmployeeNames);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)1, baPayments);
        checker.checkListCells( sourceSheet, 29, resultSheet, 51, (short)2, baBonuses);
        checker.checkRows(sourceSheet, resultSheet, 31, 30, 1, true);
        sourceSheet = sourceWorkbook.getSheetAt( 1 );
        resultSheet = resultWorkbook.getSheetAt( 1 );
        assertEquals("Number of rows on Sheet 2 is not correct", 1, resultSheet.getLastRowNum() + 1);
        checker.setIgnoreFirstLastCellNums( true );
        checker.checkRows( sourceSheet, resultSheet, 11, 0, 1, true);
        is.close();
    }

    public void testForIfTagMergeCellsExport() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTagMergeXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        // TODO: need to check the result workbook is correct
        is.close();
    }

    public void testForIfTagOneRowExport() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        saveWorkbook(resultWorkbook, forifTagOneRowDestXLS);
        is.close();

        is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRowXLS));
        Workbook sourceWorkbook = WorkbookFactory.create(is);
        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkbook.getSheetAt(0);

        Map props = new HashMap();
        props.put( "${department.name}", "IT");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 1, true);

        for(int i = 0; i < itEmployeeNames.length; i++){
            props.put("${employee.name}", itEmployeeNames[i]);
            props.put("${employee.payment}", itPayments[i]);
            props.put("${employee.bonus}", itBonuses[i]);
            short srcCol = 7;
            if( itPayments[i].doubleValue() > 2000 ){
                srcCol = 4;
            }
            checker.checkCells(sourceSheet, resultSheet, 2, (short)2, 1, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 2, srcCol, 1, (short)(i*2 + 2), false);
            checker.checkCells(sourceSheet, resultSheet, 3, (short)2, 2, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 3, srcCol, 2, (short)(i*2 + 2), false);
        }

        for(int i = 0; i < hrEmployeeNames.length; i++){
            props.put("${employee.name}", hrEmployeeNames[i]);
            props.put("${employee.payment}", hrPayments[i]);
            props.put("${employee.bonus}", hrBonuses[i]);
            short srcCol = 7;
            if( hrPayments[i].doubleValue() > 2000 ){
                srcCol = 4;
            }
            checker.checkCells(sourceSheet, resultSheet, 2, (short)2, 4, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 2, srcCol, 4, (short)(i*2 + 2), false);
            checker.checkCells(sourceSheet, resultSheet, 3, (short)2, 5, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 3, srcCol, 5, (short)(i*2 + 2), false);
        }

        for(int i = 0; i < baEmployeeNames.length; i++){
            props.put("${employee.name}", baEmployeeNames[i]);
            props.put("${employee.payment}", baPayments[i]);
            props.put("${employee.bonus}", baBonuses[i]);
            short srcCol = 7;
            if( baPayments[i].doubleValue() > 2000 ){
                srcCol = 4;
            }
            checker.checkCells(sourceSheet, resultSheet, 2, (short)2, 7, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 2, srcCol, 7, (short)(i*2 + 2), false);
            checker.checkCells(sourceSheet, resultSheet, 3, (short)2, 8, (short)(i*2 + 1), false);
            checker.checkCells(sourceSheet, resultSheet, 3, srcCol, 8, (short)(i*2 + 2), false);
        }

        is.close();
    }

    public void testForGroupBy() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        List deps = new ArrayList( departments );
        // adding department with null values to check grouping with null values
        deps.add(mgrDepartment);
        beans.put( "departments", deps );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        is.close();
        ((Department)departments.get(0)).getStaff().clear();
        is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
    }

    public void testForEachSelectWhenConditionIsNotMet() throws IOException, InvalidFormatException {
        Map beans = new HashMap();
        List employees = itDepartment.getStaff();
        beans.put("employees", employees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream("/templates/select2.xls"));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        Sheet sheet = resultWorkbook.getSheetAt(0);
        assertEquals( "Number of rows is incorrect", 1, sheet.getLastRowNum());
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        String empName = cell.getRichStringCellValue().getString();
        assertEquals("Cell value is incorrect", "Last line", empName);
        is.close();

    }


    public void testForEachSelect() throws IOException, InvalidFormatException {
        Map beans = new HashMap();
        String[] selectedEmployees = new String[]{"Oleg", "Neil", "John"};
        List employees = itDepartment.getStaff();
        beans.put("employees", employees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(selectXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        Sheet sheet = resultWorkbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        for(int i = 0; i < selectedEmployees.length; i++){
            Cell cell = row.getCell(i);
            String empName = cell.getRichStringCellValue().getString();
            assertEquals("Selected employees are incorrect", selectedEmployees[i], empName);
        }
        is.close();
    }

    public void testOutTagInOneRow() throws IOException, InvalidFormatException {
        Map beans = new HashMap();
        List employees = itDepartment.getStaff();
        beans.put("employees", employees);
        beans.put("emp", employees.get(0));
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(outTagOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        transformer.setJexlInnerCollectionsAccess(true);
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        Sheet sheet = resultWorkbook.getSheetAt(0);
        int index = 0;
        for (int i = 0; i < employees.size(); i++) {
            Employee employee = (Employee) employees.get(i);
            if( employee.getPayment().doubleValue() > 2000 ){
                Row row = sheet.getRow(index);
                index++;
                assertNotNull("Row must not be null", row);
                assertEquals("Employee names are not equal", employee.getName(), row.getCell(0).getRichStringCellValue().getString());
                assertEquals("Employee payments are not equal", employee.getPayment().doubleValue(), row.getCell(1).getNumericCellValue(), 1e-6);
                assertEquals("Employee bonuses are not equal", employee.getBonus().doubleValue(), row.getCell(2).getNumericCellValue(), 1e-6);
            }
        }
        Row row = sheet.getRow( index );
        Employee employee = (Employee) employees.get(0);
        assertEquals("Employee names are not equal", employee.getName(), row.getCell(0).getRichStringCellValue().getString());
        assertEquals("Employee payments are not equal", employee.getPayment().doubleValue(), row.getCell(1).getNumericCellValue(), 1e-6);
        assertEquals("Employee bonuses are not equal", employee.getBonus().doubleValue(), row.getCell(2).getNumericCellValue(), 1e-6);
        is.close();
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


    public void testForOneRow() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
//        beans.put( "departments", departments );
        beans.put( "itDep", itDepartment );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        Object[] values = new Object[]{"IT", "IT", null, "Elsa", new Double(1500), "Oleg", new Double(2300),
                "Neil", new Double(2500), "Maria", new Double(1700), "John", new Double(2800), "IT", "IT", "IT"};
        checker.checkRow(resultSheet, 0, 0, 13, values);
        saveWorkbook(resultWorkbook, forOneRowDestXLS);
    }

    public void testDoubleForEachInOneRow() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "itDep", itDepartment );
        beans.put( "mgrDep", mgrDepartment );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(doubleForEachOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        //TODO fix this (test fails)
//        CellsChecker checker = new CellsChecker();
//        Object[] values = new Object[]{"IT", "Elsa", new Double(1500), "Oleg", new Double(2300),
//                "Neil", new Double(2500), "Maria", new Double(1700), "John", new Double(2800), "IT", "IT", "IT"};
//        checker.checkRow(resultSheet, 0, 0, 13, values);
        saveWorkbook(resultWorkbook, doubleForEachOneRowDestXLS);
    }

    public void testForOneRowMerge() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "mgrDep", mgrDepartment );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forOneRowMergeXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        Object[] values = new Object[]{"Sean", null, null, "John", null, new Double(6000), "Joerg", null, null, "MGR", null, null};
        checker.checkRow(resultSheet, 0, 0, 11, values);
        saveWorkbook(resultWorkbook, forOneRowMergeDestXLS);
    }

    public void testForOneRowMerge2() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        beans.put( "itDep", itDepartment );
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forOneRowMerge2XLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        Object[] values = new Object[]{"Elsa", null, new Double(1500), "Oleg", null, new Double(2300), "Neil", null, new Double(2500),
                "Maria", null, new Double(1700), "John", null, new Double(2800), "IT", null, null};
        checker.checkRow(resultSheet, 0, 0, values.length - 1, values);
        saveWorkbook(resultWorkbook, forOneRowMerge2DestXLS);
    }

    //TODO: finish test
    public void testForGroup() throws InvalidFormatException, IOException {
        Map<String, Object> beans = new HashMap<String, Object>();
        prepareData(beans);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping4XLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkRow(resultSheet, 26, 0, 0, new Object[]{"Date: 01 May 2011"});
        checker.checkRow(resultSheet, 37, 0, 6, new Object[]{"01 May 2011", new Integer(0),new Integer(0),new Integer(1),new Integer(0),new Integer(0),new Integer(2) });
        saveWorkbook(resultWorkbook, grouping4DestXLS);
    }

    private void prepareData(Map<String, Object> beans){
        List<Feedback> feedbackForDay = new ArrayList<Feedback>();
        List<FeedbackCount> feedbackCountForDay = new ArrayList<FeedbackCount>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMM yyyy", Locale.US);
        beans.put("user", "Test user");
        Calendar cal = Calendar.getInstance();
        cal.set(2011, 3, 29);
        beans.put("dateFormat", dateFormat);
        beans.put("start", cal.getTime());
        cal.set(2011, 4, 3);
        beans.put("end", cal.getTime());

        cal.set(2011, 3, 29);
        Feedback feed = new Feedback(5, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);
        cal.set(2011, 3, 31);
        feed = new Feedback(5, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);
        cal.set(2011, 4, 1);
        feed = new Feedback(2, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);
        feed = new Feedback(5, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);
        feed = new Feedback(5, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);
        cal.set(2011, 4, 3);
        feed = new Feedback(0, cal.getTime(), "Well Done", "Graham Rhodes", "Test user");
        feedbackForDay.add(feed);

        Collections.sort(feedbackForDay);
        beans.put("feedback", feedbackForDay);

        FeedbackCount count = new FeedbackCount();
        cal.set(2011, 3, 29);
        count.setDate(cal.getTime());
        count.setStar0(0);
        count.setStar1(0);
        count.setStar2(0);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(1);
        feedbackCountForDay.add(count);
        count = new FeedbackCount();
        cal.set(2011, 3, 30);
        count.setDate(cal.getTime());
        count.setStar0(0);
        count.setStar1(0);
        count.setStar2(0);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(0);
        feedbackCountForDay.add(count);
        count = new FeedbackCount();
        cal.set(2011, 3, 31);
        count.setDate(cal.getTime());
        count.setStar0(0);
        count.setStar1(0);
        count.setStar2(0);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(1);
        feedbackCountForDay.add(count);
        count = new FeedbackCount();
        cal.set(2011, 4, 1);
        count.setDate(cal.getTime());
        count.setStar0(0);
        count.setStar1(0);
        count.setStar2(1);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(2);
        feedbackCountForDay.add(count);
        count = new FeedbackCount();
        cal.set(2011, 4, 2);
        count.setDate(cal.getTime());
        count.setStar0(0);
        count.setStar1(0);
        count.setStar2(0);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(566);
        feedbackCountForDay.add(count);
        count = new FeedbackCount();
        cal.set(2011, 4, 3);
        count.setDate(cal.getTime());
        count.setStar0(1);
        count.setStar1(0);
        count.setStar2(0);
        count.setStar3(0);
        count.setStar4(0);
        count.setStar5(123);
        feedbackCountForDay.add(count);
        beans.put("feedbackCounts", feedbackCountForDay);
    }

    public void testArrayProcessing() throws IOException, ParsePropertyException, InvalidFormatException {
        Map beans = new HashMap();
        int[] arr = {1,2,3};
        beans.put( "arr", arr  );
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(arrayXLS));
        XLSTransformer transformer = new XLSTransformer();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        Sheet resultSheet = resultWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkCell( resultSheet, 0, 0, 1);
        checker.checkCell( resultSheet, 1, 0, 2);
        checker.checkCell( resultSheet, 2, 0, 3);
        saveWorkbook(resultWorkbook, arrayDestXLS);
    }

}

