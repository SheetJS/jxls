package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.bean.*;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author Leonid Vysochyn
 */
public class XLSTransformerTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());

    public static final String simpleBeanXLS = "/templates/simplebean.xls";
    public static final String simpeBeanDestXLS = "target/simplebean_output.xls";

    public static final String hideSheetsXLS = "/templates/hidesheets.xls";
    public static final String hideSheetsDestXLS = "target/hidesheets_output.xls";

    public static final String beanWithListXLS = "/templates/beanwithlist.xls";
    public static final String beanWithListDestXLS = "target/beanwithlist_output.xls";

    public static final String emptyBeansXLS = "/templates/beanwithlist.xls";
    public static final String emptyBeansDestXLS = "target/emptybeans_output.xls";

    public static final String formulasXLS = "/templates/formulas.xls";
    public static final String formulasDestXLS = "target/formulas_output.xls";
    public static final String formulas2XLS = "/templates/formulas2.xls";
    public static final String formulas2DestXLS = "target/formulas2_output.xls";

    public static final String multipleListRowsXLS = "/templates/multiplelistrows.xls";
    public static final String multipleListRowsDestXLS = "target/multiplelistrows_output.xls";

    public static final String grouping1XLS = "/templates/grouping1.xls";
    public static final String grouping1DestXLS = "target/grouping1_output.xls";

    public static final String groupingFormulasXLS = "/templates/groupingformulas.xls";
    public static final String groupingFormulasDestXLS = "target/groupingformulas_output.xls";

    public static final String grouping2XLS = "/templates/grouping2.xls";
    public static final String grouping2DestXLS = "target/grouping2_output.xls";

    public static final String grouping3XLS = "/templates/grouping3.xls";
    public static final String grouping3DestXLS = "target/grouping3_output.xls";

    public static final String mergeCellsListXLS = "/templates/mergecellslist.xls";
    public static final String mergeCellsListDestXLS = "target/mergecellslist_output.xls";

    public static final String mergeMultipleListRowsXLS = "/templates/mergemultiplelistrows.xls";
    public static final String mergeMultipleListRowsDestXLS = "target/mergemultiplelistrows_output.xls";

    public static final String severalPropertiesInCellXLS = "/templates/severalpropertiesincell.xls";
    public static final String severalPropertiesInCellDestXLS = "target/severalpropertiesincell_output.xls";

    public static final String parallelTablesXLS = "/templates/paralleltables.xls";
    public static final String parallelTablesDestXLS = "target/paralleltables_output.xls";

    public static final String severalListsInRowXLS = "/templates/severallistsinrow.xls";
    public static final String severalListsInRowDestXLS = "target/severallistsinrow_output.xls";

    public static final String fixedSizeListXLS = "/templates/fixedsizelist.xls";
    public static final String fixedSizeListDestXLS = "target/fixedsizelist_output.xls";

    public static final String expressions1XLS = "/templates/expressions1.xls";
    public static final String expressions1DestXLS = "target/expressions1_output.xls";

    public static final String iftagXLS = "/templates/iftag.xls";
    public static final String iftagDestXLS = "target/iftag_output.xls";

    public static final String forifTag2XLS = "/templates/foriftag2.xls";
    public static final String forifTag2DestXLS = "target/foriftag2_output.xls";

    public static final String poiobjectsXLS = "/templates/poiobjects.xls";
    public static final String poiobjectsDestXLS = "target/poiobjects_output.xls";

    public static final String forifTag3XLS = "/templates/foriftag3.xls";
    public static final String forifTag3DestXLS = "target/foriftag3_output.xls";
    
    public static final String forifTag3OutTagXLS = "/templates/foriftag3OutTag.xls";
    public static final String forifTag3OutTagDestXLS = "target/foriftag3OutTag_output.xls";

    public static final String forifTagMergeXLS = "/templates/foriftagmerge.xls";
    public static final String forifTagMergeDestXLS = "target/foriftagmerge_output.xls";

    public static final String employeeNotesXLS = "/templates/employeeNotes.xls";
    public static final String employeeNotesDestXLS = "target/employeeNotes_output.xls";
    public static final String employeeNotesRusDestXLS = "target/employeeNotesRus_output.xls";

    public static final String forifTagOneRowXLS = "/templates/foriftagOneRow.xls";
    public static final String forifTagOneRowDestXLS = "target/foriftagOneRow_output.xls";

    public static final String dynamicColumnsXLS = "/templates/dynamicColumns.xls";
    public static final String dynamicColumnsDestXLS = "target/dynamicColumns_output.xls";

    public static final String forifTagOneRow2XLS = "/templates/foriftagOneRow2.xls";
    public static final String forifTagOneRowDest2XLS = "target/foriftagOneRow2_output.xls";

    public static final String multipleSheetListXLS = "/templates/multipleSheetList.xls";
    public static final String multipleSheetListDestXLS = "target/multipleSheetList_output.xls";

    public static final String multipleSheetList2XLS = "/templates/multipleSheetList2.xls";
    public static final String multipleSheetList2DestXLS = "target/multipleSheetList2_output.xls";

    public static final String multiTabXLS = "/templates/multi-tab-template.xls";
    public static final String multiTabDestXLS = "target/multi-tab_output.xls";

    public static final String groupTagXLS = "/templates/groupTag.xls";
    public static final String groupTagDestXLS = "target/groupTag_output.xls";

    public static final String jexlXLS = "/templates/jexl.xls";
    public static final String jexlDestXLS = "target/jexl_output.xls";

    public static final String forGroupByXLS = "/templates/forgroup.xls";
    public static final String forGroupByDestXLS = "target/forgroup_output.xls";


    SimpleBean simpleBean1;
    SimpleBean simpleBean2;
    SimpleBean simpleBean3;
    BeanWithList beanWithList;
    List beans = new ArrayList();
    List itEmployees = new ArrayList();
    BeanWithList listBean1 = new BeanWithList("List bean 1");
    BeanWithList listBean2 = new BeanWithList("List bean 2");

    Object[] names = new Object[]{"Bean 1", "Bean 2", "Bean 3"};
    Object[] doubleValues = new Object[]{new Double(100.34567), new Double(555.3), new Double(777.569)};
    Object[] intValues = new Object[]{new Integer(10), new Integer(123), new Integer(10234)};
    Object[] dateValues = new Object[]{new Date(), null, new Date()};

    Object[] names2 = new Object[]{"Bean 1", "Bean 2", "Bean 3", "Bean 4", "Bean 5", "Bean 6", "Bean 7"};
    Object[] doubleValues2 = new Object[]{new Double(111.222), new Double(222.333), new Double(333.444),
            new Double(444.555), new Double(555.666), new Double(666.777), new Double(777.888)};
    Object[] intValues2 = new Object[]{new Integer(11), new Integer(12), new Integer(13), new Integer(14), new Integer(15),
            new Integer(16), new Integer(17)};


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

    int[] amounts = {1, 2, 4, 6, 7, 8, 9, 10, 11, 13, 15, 18, 20, 21, 22};
    List amountBeans = new ArrayList();

    public XLSTransformerTest() {
    }

    public XLSTransformerTest(String s) {
        super(s);
    }

    protected void setUp() throws Exception {
        super.setUp();
        simpleBean1 = new SimpleBean(names[0].toString(), (Double) doubleValues[0], (Integer) intValues[0], (Date) dateValues[0]);
        simpleBean2 = new SimpleBean(names[1].toString(), (Double) doubleValues[1], (Integer) intValues[1], (Date) dateValues[1]);
        simpleBean3 = new SimpleBean(names[2].toString(), (Double) doubleValues[2], (Integer) intValues[2], (Date) dateValues[2]);

        listBean2.addBean( new SimpleBean(names2[0].toString(), (Double) doubleValues2[0], (Integer) intValues2[0]) );
        listBean2.addBean( new SimpleBean(names2[1].toString(), (Double) doubleValues2[1], (Integer) intValues2[1]) );
        listBean2.addBean( new SimpleBean(names2[2].toString(), (Double) doubleValues2[2], (Integer) intValues2[2]) );
        listBean2.addBean( new SimpleBean(names2[3].toString(), (Double) doubleValues2[3], (Integer) intValues2[3]) );
        listBean2.addBean( new SimpleBean(names2[4].toString(), (Double) doubleValues2[4], (Integer) intValues2[4]) );
        listBean2.addBean( new SimpleBean(names2[5].toString(), (Double) doubleValues2[5], (Integer) intValues2[5]) );
        listBean2.addBean( new SimpleBean(names2[6].toString(), (Double) doubleValues2[6], (Integer) intValues2[6]) );

        simpleBean1.setOther(simpleBean2);
        simpleBean2.setOther(simpleBean3);
//        simpleBean3.setOther( simpleBean1 );

        beanWithList = new BeanWithList("Bean With List", new Double(1976.1202));


        beans.add(simpleBean1);
        beans.add(simpleBean2);
        beans.add(simpleBean3);

        listBean1.addBean( simpleBean1 );
        listBean1.addBean( simpleBean2 );
        listBean1.addBean( simpleBean3 );

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


        beanWithList.setBeans(beans);

        propertyMap.put("${bean.name}", simpleBean1.getName());
        propertyMap.put("${bean.doubleValue}", simpleBean1.getDoubleValue());
        propertyMap.put("${bean.intValue}", simpleBean1.getIntValue());
        propertyMap.put("${bean.dateValue}", simpleBean1.getDateValue());
        propertyMap.put("${bean.other.name}", simpleBean1.getOther().getName());
        propertyMap.put("${bean.other.intValue}", simpleBean1.getOther().getIntValue());
        propertyMap.put("${bean.other.doubleValue}", simpleBean1.getOther().getDoubleValue());
        propertyMap.put("${bean.other.dateValue}", simpleBean1.getOther().getDateValue());

        propertyMap.put("${listBean.name}", beanWithList.getName());
//        propertyMap.put("${listBean.beans.name}", beanWithList.getBeans());

        for (int i = 0; i < amounts.length; i++) {
            int amount = amounts[i];
            amountBeans.add( new SimpleBean( amount ) );
        }

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

    Map propertyMap = new HashMap();

    public void testSimpleBeanExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("bean", simpleBean1);
        Calendar calendar = Calendar.getInstance();
        calendar.set(2006, 8, 19);
        beans.put("calendar", calendar);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(simpleBeanXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Numbers differ in source and result sheets", sourceSheet.getLastRowNum(), resultSheet.getLastRowNum());
        CellsChecker checker = new CellsChecker(propertyMap);
        propertyMap.put("${calendar}", calendar);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 6);

        is.close();
        saveWorkbook(resultWorkbook, simpeBeanDestXLS);
    }

    public void testBeanWithListExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        beans.put("beans", beanWithList.getBeans());
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(beanWithListXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(beanWithListXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + beanWithList.getBeans().size() - 1, resultSheet.getLastRowNum());

        Map listPropMap = new HashMap();
        listPropMap.put("${listBean.name}", beanWithList.getName());
        CellsChecker checker = new CellsChecker(listPropMap);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 0, names);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 1, doubleValues);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, new Object[]{new Integer(123), new Integer(10234), null});
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 3, dateValues);
        is.close();
        saveWorkbook(resultWorkbook, beanWithListDestXLS);
    }

    public void testFormulas2() throws IOException {
        Map beans = new HashMap();
        beans.put("departments", departments);

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formulas2XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();

        saveWorkbook(resultWorkbook, formulasDestXLS);

    }

    public void testFormulas() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        beans.put("departments", departments);
        beans.put( "t1", amountBeans );
        //todo comment this line to work on #VALUE! formula cell problem
//        simpleBean3.setOther( simpleBean1 );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(formulasXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(formulasXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + beanWithList.getBeans().size() - 1, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${listBean.name}", beanWithList.getName());
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 0, names);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 1, doubleValues);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, new Object[]{new Integer(123), new Integer(10234)});
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 3, dateValues);
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 6, (short) 1, "SUM(B4:B6)");
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 6, (short) 2, "SUM(C4:C6)");
        checker.checkFormulaCell(sourceSheet, 6, resultSheet, 8, (short) 1, "MAX(B7,C7)");

        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 3, (short) 4, "B4+C4");
        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 4, (short) 4, "B5+C5");
        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 5, (short) 4, "B6+C6");

        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 6, (short) 4, "SUM(E4:E6)");

        checker.checkFormulaCell( sourceSheet, 8, resultSheet, 10, (short) 1, "SUM(B4:B6)");
        checker.checkFormulaCell( sourceSheet, 8, resultSheet, 10, (short) 2, "SUM(C4:C6)");
        checker.checkFormulaCell( sourceSheet, 8, resultSheet, 10, (short) 4, "SUM(E4:E6)");
        checker.checkFormulaCell(sourceSheet, 10, resultSheet, 12, (short) 1, "MAX(B7,C7)");

        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 23, (short)1, "SUM(B19:B23)");
        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 32, (short)1, "SUM(B29:B32)");
        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 40, (short)1, "SUM(B38:B40)");

        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 23, (short)3, "SUM(D19:D23)");
        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 32, (short)3, "SUM(D29:D32)");
        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 40, (short)3, "SUM(D38:D40)");

        checker.checkFormulaCell( sourceSheet, 22, resultSheet, 41, (short)1, "SUM(B24,B33,B41)");
        checker.checkFormulaCell( sourceSheet, 22, resultSheet, 41, (short)3, "SUM(D24,D33,D41)");

        checker.checkFormulaCell( sourceSheet, 18, resultSheet, 18, (short)3, "B19*(1+C19)");
        checker.checkFormulaCell( sourceSheet, 18, resultSheet, 22, (short)3, "B23*(1+C23)");
        checker.checkFormulaCell( sourceSheet, 18, resultSheet, 28, (short)3, "B29*(1+C29)");
        checker.checkFormulaCell( sourceSheet, 19, resultSheet, 31, (short)3, "B32*(1+C32)");
        checker.checkFormulaCell( sourceSheet, 19, resultSheet, 37, (short)3, "B38*(1+C38)");
        checker.checkFormulaCell( sourceSheet, 19, resultSheet, 39, (short)3, "B40*(1+C40)");

        checker.checkFormulaCell( sourceSheet, 24, resultSheet, 43, (short)1, "'Sheet 2'!B55");

        sourceSheet = sourceWorkbook.getSheetAt( 1 );
        resultSheet = resultWorkbook.getSheetAt( 1 );
        checker.checkFormulaCell( sourceSheet, 0, resultSheet, 0, (short)1, "SUM(Sheet1!B4:B6)");
        checker.checkFormulaCell( sourceSheet, 0, resultSheet, 0, (short)2, "SUM(Sheet1!C4:C6)");
        checker.checkFormulaCell( sourceSheet, 0, resultSheet, 0, (short)4, "SUM(Sheet1!E4:E6)");
        checker.checkFormulaCell( sourceSheet, 2, resultSheet, 2, (short)1, "MAX(Sheet1!B7,Sheet1!C7)");
        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 4, (short)1, "Sheet1!B13");

        checker.checkFormulaCell( sourceSheet, 15, resultSheet, 24, (short)1, "SUM(B10,B13,B16,B19,B22)");
        checker.checkFormulaCell( sourceSheet, 15, resultSheet, 40, (short)1, "SUM(B29,B32,B35,B38)");
        checker.checkFormulaCell( sourceSheet, 15, resultSheet, 53, (short)1, "SUM(B45,B48,B51)");

        checker.checkFormulaCell( sourceSheet, 18, resultSheet, 55, (short)1, "Sheet1!D24");
        checker.checkFormulaCell( sourceSheet, 19, resultSheet, 56, (short)1, "Sheet1!D33");
        checker.checkFormulaCell( sourceSheet, 20, resultSheet, 57, (short)1, "Sheet1!D41");
        // todo Create checks for "Sheet 3"

        is.close();
        saveWorkbook(resultWorkbook, formulasDestXLS);
    }

    public void testMultipleListRows() throws IOException, ParsePropertyException {

        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(multipleListRowsXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(multipleListRowsXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + (beanWithList.getBeans().size() - 1) * 4, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${listBean.beans.name}//:3", names[0]);
        props.put("${listBean.beans.doubleValue}", doubleValues[0]);
        props.put("${listBean.beans.other.intValue}", simpleBean1.getOther().getIntValue());
        props.put("${listBean.beans.dateValue}", dateValues[0]);
        props.put("//listBean.beans", "");
        props.put("Int Value://listBean.beans", "Int Value:");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 3, 4);

        props.clear();
        props.put("${listBean.beans.name}//:3", names[1]);
        props.put("${listBean.beans.doubleValue}", doubleValues[1]);
        props.put("${listBean.beans.other.intValue}", simpleBean2.getOther().getIntValue());
        props.put("${listBean.beans.dateValue}", dateValues[1]);
        props.put("//listBean.beans", "");
        props.put("Int Value://listBean.beans", "Int Value:");
        checker = new CellsChecker(props);

        checker.checkRows(sourceSheet, resultSheet, 3, 7, 4);

        props.clear();
        props.put("${listBean.beans.name}//:3", names[2]);
        props.put("${listBean.beans.doubleValue}", doubleValues[2]);
        props.put("${listBean.beans.other.intValue}", "");
        props.put("${listBean.beans.dateValue}", dateValues[2]);
        props.put("//listBean.beans", "");
        props.put("Int Value://listBean.beans", "Int Value:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 11, 4);

        is.close();
        saveWorkbook(resultWorkbook, multipleListRowsDestXLS);
    }

    public void testMergedMultipleListRows() throws IOException, ParsePropertyException {

        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(mergeMultipleListRowsXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(mergeMultipleListRowsXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + (beanWithList.getBeans().size() - 1) * 4, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${listBean.beans.name}//:3", names[0]);
        props.put("${listBean.beans.doubleValue}", doubleValues[0]);
        props.put("${listBean.beans.dateValue}", dateValues[0]);
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 3, 4);

        props.clear();
        props.put("${listBean.beans.name}//:3", names[1]);
        props.put("${listBean.beans.doubleValue}", doubleValues[1]);
        props.put("${listBean.beans.dateValue}", dateValues[1]);
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 7, 4);

        props.clear();
        props.put("${listBean.beans.name}//:3", names[2]);
        props.put("${listBean.beans.doubleValue}", doubleValues[2]);
        props.put("${listBean.beans.dateValue}", dateValues[2]);
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 11, 4);

        assertEquals( "Incorrect number of merged regions", 9, resultSheet.getNumMergedRegions() );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(3, (short)0, 3, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(7, (short)0, 7, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(11, (short)0, 11, (short)2) ) );

        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(4, (short)1, 4, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(8, (short)1, 8, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(12, (short)1, 12, (short)2) ) );

        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(5, (short)1, 6, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(9, (short)1, 10, (short)2) ) );
        assertTrue( "Merged Region not found", isMergedRegion( resultSheet, new Region(13, (short)1, 14, (short)2) ) );

        is.close();
        saveWorkbook(resultWorkbook, mergeMultipleListRowsDestXLS);
    }

    public void testGrouping1() throws IOException, ParsePropertyException {
        BeanWithList beanWithList2 = new BeanWithList("2nd bean with list", new Double(22.22));
        List beans2 = new ArrayList();
        beans2.add(new SimpleBean("bean 21", new Double(21.21), new Integer(21), new Date()));
        beans2.add(new SimpleBean("bean 22", new Double(22.22), new Integer(22), new Date()));
        beanWithList2.setBeans(beans2);
        BeanWithList beanWithList3 = new BeanWithList("3d bean with list", new Double(333.333));
        List beans3 = new ArrayList();
        beans3.add(new SimpleBean("bean 31", new Double(31.31), new Integer(31), new Date()));
        beans3.add(new SimpleBean("bean 32", new Double(32.32), new Integer(32), new Date()));
        beanWithList3.setBeans(beans3);
        List mainList = new ArrayList();
        mainList.add(beanWithList2);
        mainList.add(beanWithList3);
        BeanWithList bean = new BeanWithList("Root", new Double(1111.1111));
        bean.setBeans(mainList);
        Map beans = new HashMap();
        beans.put("mainBean", bean);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping1XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(grouping1XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + 6, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${mainBean.beans.name}//:3", "2nd bean with list");
        props.put("${mainBean.beans.beans.name}", "bean 21");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(21.21));
        props.put("${mainBean.name}", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 1, 3);
        props.clear();
        props.put("${mainBean.beans.beans.name}", "bean 22");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(22.22));
        props.put("${mainBean.name}", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 4, 1);
        checker.checkRows(sourceSheet, resultSheet, 4, 5, 1);
        props.clear();
        props.put("${mainBean.beans.name}//:3", "3d bean with list");
        props.put("${mainBean.beans.beans.name}", "bean 31");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(31.31));
        props.put("${mainBean.name}", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 6, 3);
        props.clear();
        props.put("${mainBean.beans.beans.name}", "bean 32");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(32.32));
        props.put("${mainBean.name}", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 9, 1);
        checker.checkRows(sourceSheet, resultSheet, 4, 10, 1);

        is.close();
        saveWorkbook(resultWorkbook, grouping1DestXLS);
    }

    public void testMergeCellsList() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(mergeCellsListXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(mergeCellsListXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + beanWithList.getBeans().size() - 1, resultSheet.getLastRowNum());

        Map listPropMap = new HashMap();
        listPropMap.put("${listBean.name}", beanWithList.getName());
        CellsChecker checker = new CellsChecker(listPropMap);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 0, names);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 1, intValues);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 3, doubleValues);
        assertEquals( "Incorrect number of merged regions", 3, resultSheet.getNumMergedRegions() );
        assertTrue( "Merged Region (3,1,3,2) not found", isMergedRegion( resultSheet, new Region(3, (short)1, 3, (short)2) ) );
        assertTrue( "Merged Region (4,1,4,2) not found", isMergedRegion( resultSheet, new Region(4, (short)1, 4, (short)2) ) );
        assertTrue( "Merged Region (5,1,5,2) not found", isMergedRegion( resultSheet, new Region(5, (short)1, 5, (short)2) ) );

        is.close();
        saveWorkbook(resultWorkbook, mergeCellsListDestXLS);
    }

    protected static boolean isMergedRegion(HSSFSheet sheet, Region region){
        for(int i = 0; i < sheet.getNumMergedRegions(); i++){
            Region mgdRegion = sheet.getMergedRegionAt( i );
            if( mgdRegion.equals( region ) ){
                return true;
            }
//            log.info("(" + mgdRegion.getRowFrom() + ", " + mgdRegion.getColumnFrom() + ", " + mgdRegion.getRowTo() + ", " + mgdRegion.getColumnTo() + ")");
        }
        return false;
    }

    public void testGrouping2() throws IOException, ParsePropertyException {
        BeanWithList beanWithList2 = new BeanWithList("2nd bean with list", new Double(22.22));
        List beans2 = new ArrayList();
        beans2.add(new SimpleBean("bean 21", new Double(21.21), new Integer(21), new Date()));
        beans2.add(new SimpleBean("bean 22", new Double(22.22), new Integer(22), new Date()));
        beanWithList2.setBeans(beans2);
        BeanWithList beanWithList3 = new BeanWithList("3d bean with list", new Double(333.333));
        List beans3 = new ArrayList();
        beans3.add(new SimpleBean("bean 31", new Double(31.31), new Integer(31), new Date()));
        beans3.add(new SimpleBean("bean 32", new Double(32.32), new Integer(32), new Date()));
        beanWithList3.setBeans(beans3);
        List mainList = new ArrayList();
        mainList.add(beanWithList2);
        mainList.add(beanWithList3);
        BeanWithList bean = new BeanWithList("Root", new Double(1111.1111));
        bean.setBeans(mainList);
        Map beans = new HashMap();
        beans.put("mainBean", bean);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping2XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(grouping2XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", 14, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${mainBean.beans.name}//:4", "2nd bean with list");
        props.put("${mainBean.beans.beans.name}//:1", "bean 21");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(21.21));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 1, 4);
        props.clear();
        props.put("${mainBean.beans.beans.name}//:1", "bean 22");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(22.22));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 5, 3);
        props.clear();
        props.put("${mainBean.beans.name}//:4", "3d bean with list");
        props.put("${mainBean.beans.beans.name}//:1", "bean 31");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(31.31));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 8, 4);
        props.clear();
        props.put("${mainBean.beans.beans.name}//:1", "bean 32");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(32.32));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 12, 3);

        is.close();
        saveWorkbook(resultWorkbook, grouping2DestXLS);
    }

    public void testGrouping3() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("departments", departments);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(grouping3XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(grouping3XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);

        Map props = new HashMap();
        props.put("${departments.name}//:4", "IT");
        props.put("Department//departments", "Department");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 0, itEmployeeNames);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 1, itPayments);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, itBonuses);
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 3, (short)3, "B4*(1+C4)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 4, (short)3, "B5*(1+C5)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 7, (short)3, "B8*(1+C8)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 8, (short)1, "SUM(B4:B8)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 6, (short)3, "SUM(D4:D8)");
        props.clear();
        props.put("${departments.name}//:4", "HR");
        props.put("Department//departments", "Department");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 9, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 3, resultSheet, 12, (short) 0, hrEmployeeNames);
        checker.checkListCells(sourceSheet, 3, resultSheet, 12, (short) 1, hrPayments);
        checker.checkListCells(sourceSheet, 3, resultSheet, 12, (short) 2, hrBonuses);
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 12, (short)3, "B13*(1+C13)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 13, (short)3, "B14*(1+C14)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 15, (short)3, "B16*(1+C16)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 16, (short)1, "SUM(B13:B16)");
//        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 16, (short)3, "SUM(D13:D16)");
        props.clear();
        props.put("${departments.name}//:4", "BA");
        props.put("Department//departments", "Department");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 17, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 3, resultSheet, 20, (short) 0, baEmployeeNames);
        checker.checkListCells(sourceSheet, 3, resultSheet, 20, (short) 1, baPayments);
        checker.checkListCells(sourceSheet, 3, resultSheet, 20, (short) 2, baBonuses);
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 20, (short)3, "B21*(1+C21)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 21, (short)3, "B22*(1+C22)");
        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 22, (short)3, "B23*(1+C23)");
        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 23, (short)1, "SUM(B21:B23)");
        checker.checkFormulaCell( sourceSheet, 4, resultSheet, 23, (short)3, "SUM(D21:D23)");

        saveWorkbook(resultWorkbook, grouping3DestXLS);
    }

    public void testGroupingFormulas() throws IOException, ParsePropertyException {
        BeanWithList beanWithList2 = new BeanWithList("2nd bean with list", new Double(22.22));
        List beans2 = new ArrayList();
        beans2.add(new SimpleBean("bean 21", new Double(21.21), new Integer(21), new Date()));
        beans2.add(new SimpleBean("bean 22", new Double(22.22), new Integer(22), new Date()));
        beanWithList2.setBeans(beans2);
        BeanWithList beanWithList3 = new BeanWithList("3d bean with list", new Double(333.333));
        List beans3 = new ArrayList();
        beans3.add(new SimpleBean("bean 31", new Double(31.31), new Integer(31), new Date()));
        beans3.add(new SimpleBean("bean 32", new Double(32.32), new Integer(32), new Date()));
        beanWithList3.setBeans(beans3);
        List mainList = new ArrayList();
        mainList.add(beanWithList2);
        mainList.add(beanWithList3);
        BeanWithList bean = new BeanWithList("Root", new Double(1111.1111));
        bean.setBeans(mainList);
        Map beans = new HashMap();
        beans.put("mainBean", bean);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(groupingFormulasXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(groupingFormulasXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);
        is.close();
//        HSSFWorkbook resultWorkbook = new HSSFWorkbook( new POIFSFileSystem( new BufferedInputStream(getClass().getResourceAsStream(groupingFormulasDestXLS))));
        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + 6, resultSheet.getLastRowNum());

        Map props = new HashMap();
        props.put("${mainBean.beans.name}//:3", "2nd bean with list");
        props.put("${mainBean.beans.beans.name}", "bean 21");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(21.21));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 1, 3);
        props.clear();
        props.put("${mainBean.beans.beans.name}", "bean 22");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(22.22));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 4, 1);
//        Todo: next check requires investigation
//        Next check currently fails. It seems POI does not get the value of this formula cell correctly.
//        It returns "SUM(B9:B10)" instead of "SUM(B4:B5)". But in the output XLS file the formula is correct.
//        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 5, (short)1, "SUM(B4:B5)");
        props.clear();
        props.put("${mainBean.beans.name}//:3", "3d bean with list");
        props.put("${mainBean.beans.beans.name}", "bean 31");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(31.31));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 6, 3);
        props.clear();
        props.put("${mainBean.beans.beans.name}", "bean 32");
        props.put("${mainBean.beans.beans.doubleValue}", new Double(32.32));
        props.put("${mainBean.name}//mainBean.beans.beans", bean.getName());
        props.put("Name://mainBean.beans", "Name:");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 3, 9, 1);
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 10, (short) 1, "SUM(B9:B10)");

        saveWorkbook(resultWorkbook, groupingFormulasDestXLS);
    }

    public void testSeveralPropertiesInCell() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("bean", simpleBean1);
        beans.put("listBean", beanWithList);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(severalPropertiesInCellXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(severalPropertiesInCellXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + beanWithList.getBeans().size() - 1, resultSheet.getLastRowNum());
        Map props = new HashMap();
        props.put("Name: ${bean.name}", "Name: " + simpleBean1.getName());
        props.put("${bean.other.name} - ${bean.doubleValue},${bean.other.intValue}", simpleBean1.getOther().getName() +
                " - " + simpleBean1.getDoubleValue() + "," + simpleBean1.getOther().getIntValue());
        props.put("${bean.dateValue}", simpleBean1.getDateValue());
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum(), 6);

        Map listPropMap = new HashMap();
        listPropMap.put("[${listBean.beans.name}]", "[" + beanWithList.getName() + "]");
        checker = new CellsChecker(listPropMap);
        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 0,
                new String[]{ "[" +((SimpleBean)beanWithList.getBeans().get(0)).getName() + "]",
                        "[" + ((SimpleBean)beanWithList.getBeans().get(1)).getName() + "]",
                        "[" + ((SimpleBean)beanWithList.getBeans().get(2)).getName() + "]" });
        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 1,
                new String[]{((SimpleBean)beanWithList.getBeans().get(0)).getDoubleValue() + " yeah",
                        ((SimpleBean)beanWithList.getBeans().get(1)).getDoubleValue() + " yeah",
                        ((SimpleBean)beanWithList.getBeans().get(2)).getDoubleValue() + " yeah" });

        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 2,
                new String[]{((SimpleBean)beanWithList.getBeans().get(0)).getName() + " : " + ((SimpleBean)beanWithList.getBeans().get(0)).getDoubleValue() + "!",
                        ((SimpleBean)beanWithList.getBeans().get(1)).getName() + " : " + ((SimpleBean)beanWithList.getBeans().get(1)).getDoubleValue() + "!",
                        ((SimpleBean)beanWithList.getBeans().get(2)).getName() + " : " + ((SimpleBean)beanWithList.getBeans().get(2)).getDoubleValue() + "!" });

        is.close();
        saveWorkbook( resultWorkbook, severalPropertiesInCellDestXLS);
    }

    public void testParallelTablesExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("listBean", beanWithList);
        beans.put("bean", simpleBean2);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(parallelTablesXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(parallelTablesXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", 11, resultSheet.getLastRowNum());

        Map listPropMap = new HashMap();
        listPropMap.put("${listBean.name}", beanWithList.getName());
        listPropMap.put("Name: ${bean.name}", "Name: " + simpleBean2.getName());
        listPropMap.put("${bean.doubleValue}", simpleBean2.getDoubleValue() );
        listPropMap.put("Merged - ${bean.intValue}", "Merged - " + simpleBean2.getIntValue() );
        listPropMap.put("${bean.intValue}", simpleBean2.getIntValue() );

        CellsChecker checker = new CellsChecker(listPropMap);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);

        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, names);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 3, doubleValues);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 5, intValues);

        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 3, (short) 4, "D4+F4");
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 4, (short) 4, "D5+F5", true);
        checker.checkFormulaCell(sourceSheet, 5, resultSheet, 5, (short) 4, "D6+F6", true);

        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)0, (short)1, 7, true, true);
        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)6, (short)7, 14, true, true);

        assertEquals( "Incorrect number of merged regions", 2, resultSheet.getNumMergedRegions() );
        assertTrue( "Merged Region (4,0,4,1) not found", isMergedRegion( resultSheet, new Region(4, (short)0, 4, (short)1) ) );
        assertTrue( "Merged Region (3,6,3,7) not found", isMergedRegion( resultSheet, new Region(3, (short)6, 3, (short)7) ) );

        is.close();
        saveWorkbook( resultWorkbook, parallelTablesDestXLS);
    }

    public void testSeveralListsInRowExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("list1", listBean1);
        beans.put("list2", listBean2);
        beans.put("bean", simpleBean2);
        beans.put("staticBean", simpleBean1);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(severalListsInRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(severalListsInRowXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", 10, resultSheet.getLastRowNum());

        Map listPropMap = new HashMap();
        listPropMap.put("Name: ${list1.name}", "Name: " + listBean1.getName());
        listPropMap.put("Name: ${list2.name}", "Name: " + listBean2.getName());
        // static tables check
        listPropMap.put("Name: ${bean.name}", "Name: " + simpleBean2.getName());
        listPropMap.put("${bean.doubleValue}", simpleBean2.getDoubleValue() );
        listPropMap.put("Merged - ${bean.intValue}", "Merged - " + simpleBean2.getIntValue() );
        listPropMap.put("${bean.intValue}", simpleBean2.getIntValue() );
        listPropMap.put("Name: ${staticBean.name}", "Name: " + simpleBean1.getName() );
        listPropMap.put("${staticBean.intValue}", simpleBean1.getIntValue() );
        listPropMap.put("${staticBean.doubleValue}", simpleBean1.getDoubleValue() );

        CellsChecker checker = new CellsChecker(listPropMap);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);

        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)0, (short)1, 7, true, true);
        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)7, (short)8, 8, true, true);
        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)13, (short)14, 10, true, true);

        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, new String[]{"Name: " + names[0], "Name: " + names[1], "Name: " + names[2]} );
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 3, new String[]{names[0] + " - " + names[0] + " : " + intValues[0],
                names[1] + " - " + names[1] + " : " + intValues[1],names[2] + " - " + names[2] + " : " + intValues[2] });
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 5, doubleValues);
        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 12, intValues);
//
        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 3, (short) 6, "F4+M4");
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 4, (short) 6, "F5+M5", true);
        checker.checkFormulaCell(sourceSheet, 5, resultSheet, 5, (short) 6, "F6+M6", true);

        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 6, (short) 5, "SUM(F4:F6)");

        checker.checkListCells( sourceSheet, 3, resultSheet, 3, (short)10, doubleValues2);
        checker.checkListCells( sourceSheet, 3, resultSheet, 3, (short)11, intValues2);

        checker.checkFormulaCell(sourceSheet, 3, resultSheet, 3, (short) 9, "K4+L4");
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 4, (short) 9, "K5+L5", true);
        checker.checkFormulaCell(sourceSheet, 5, resultSheet, 5, (short) 9, "K6+L6", true);

        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 10, (short) 9, "SUM(J4:J10)");
        checker.checkFormulaCell(sourceSheet, 4, resultSheet, 10, (short) 11, "SUM(L4:L10)");

        assertEquals( "Incorrect number of merged regions", 6, resultSheet.getNumMergedRegions() );
        assertTrue( "Merged Region (4,0,4,1) not found", isMergedRegion( resultSheet, new Region(4, (short)0, 4, (short)1) ) );
        assertTrue( "Merged Region (3,7,3,8) not found", isMergedRegion( resultSheet, new Region(3, (short)7, 3, (short)8) ) );
        assertTrue( "Merged Region (3,13,3,14) not found", isMergedRegion( resultSheet, new Region(3, (short)13, 3, (short)14) ) );

        assertTrue( "Merged Region (3,3,3,4) not found", isMergedRegion( resultSheet, new Region(3, (short)3, 3, (short)4) ) );
        assertTrue( "Merged Region (4,3,4,4) not found", isMergedRegion( resultSheet, new Region(4, (short)3, 4, (short)4) ) );
        assertTrue( "Merged Region (5,3,5,4) not found", isMergedRegion( resultSheet, new Region(5, (short)3, 5, (short)4) ) );


        is.close();
        saveWorkbook( resultWorkbook, severalListsInRowDestXLS );
    }

    public void testFixedSizeCollections() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("employee", itEmployees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(fixedSizeListXLS));
        XLSTransformer transformer = new XLSTransformer();
        transformer.markAsFixedSizeCollection( "employee" );
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(fixedSizeListXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum(), resultSheet.getLastRowNum());

        Map props = new HashMap();
        CellsChecker checker;
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 2);

        checker.checkFixedListCells( sourceSheet, 2, resultSheet, 2, (short)0, itEmployeeNames);
        checker.checkFixedListCells( sourceSheet, 2, resultSheet, 2, (short)1, itPayments);
        checker.checkFixedListCells( sourceSheet, 2, resultSheet, 2, (short)2, itBonuses);

        is.close();
        saveWorkbook(resultWorkbook, fixedSizeListDestXLS);

    }

    public void testExpressions1() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("bean", simpleBean1);
        beans.put("listBean", beanWithList);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(expressions1XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(expressions1XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum() + beanWithList.getBeans().size() - 1, resultSheet.getLastRowNum());
        Map props = new HashMap();
        props.put("Name: ${bean.name}", "Name: " + simpleBean1.getName());
        props.put("${bean.other.name} - ${bean.doubleValue*2},${(bean.other.intValue + bean.doubleValue)/0.5}",
                simpleBean1.getOther().getName() +
                " - " + simpleBean1.getDoubleValue().doubleValue()*2 + "," + (simpleBean1.getOther().getIntValue().intValue() + simpleBean1.getDoubleValue().doubleValue())/0.5);
        props.put("${10*bean.doubleValue + 2.55}", new Double(simpleBean1.getDoubleValue().doubleValue() * 10 + 2.55));
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum(), 6);

        Map listPropMap = new HashMap();
        listPropMap.put("[${listBean.beans.name}]", "[" + beanWithList.getName() + "]");
        checker = new CellsChecker(listPropMap);
        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 0,
                new String[]{ "[" +((SimpleBean)beanWithList.getBeans().get(0)).getName() + "]",
                        "[" + ((SimpleBean)beanWithList.getBeans().get(1)).getName() + "]",
                        "[" + ((SimpleBean)beanWithList.getBeans().get(2)).getName() + "]" });
        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 1,
                new String[]{(((SimpleBean)beanWithList.getBeans().get(0)).getDoubleValue().doubleValue()*10.2)/10 + 1.567 + " yeah",
                        (((SimpleBean)beanWithList.getBeans().get(1)).getDoubleValue().doubleValue()*10.2)/10 + 1.567 + " yeah",
                        (((SimpleBean)beanWithList.getBeans().get(2)).getDoubleValue().doubleValue()*10.2)/10 + 1.567 + " yeah" });

        checker.checkListCells(sourceSheet, 6, resultSheet, 6, (short) 2,
                new String[]{((SimpleBean)beanWithList.getBeans().get(0)).getDoubleValue().doubleValue()  + ((SimpleBean)beanWithList.getBeans().get(0)).getIntValue().intValue()*2.1
                        + " - " + (((SimpleBean)beanWithList.getBeans().get(0)).getIntValue().intValue()*(10 + 1.1)),
                        ((SimpleBean)beanWithList.getBeans().get(1)).getDoubleValue().doubleValue()  + ((SimpleBean)beanWithList.getBeans().get(1)).getIntValue().intValue()*2.1
                        + " - " + (((SimpleBean)beanWithList.getBeans().get(1)).getIntValue().intValue()*(10 + 1.1)),
                        ((SimpleBean)beanWithList.getBeans().get(2)).getDoubleValue().doubleValue()  + ((SimpleBean)beanWithList.getBeans().get(2)).getIntValue().intValue()*2.1
                        + " - " + (((SimpleBean)beanWithList.getBeans().get(2)).getIntValue().intValue()*(10 + 1.1))});


        is.close();
        saveWorkbook( resultWorkbook, expressions1DestXLS);
    }

    public void testIfTag() throws IOException, ParsePropertyException {
        Map beans = new HashMap();

        SimpleBean simpleBean1 = new SimpleBean(names[0].toString(), (Double) doubleValues[0], (Integer) intValues[0], (Date) dateValues[0]);
        SimpleBean simpleBean2 = new SimpleBean(names[1].toString(), (Double) doubleValues[1], (Integer) intValues[1], (Date) dateValues[1]);
        SimpleBean simpleBean3 = new SimpleBean(names[2].toString(), (Double) doubleValues[2], (Integer) intValues[2], (Date) dateValues[2]);

        BeanWithList listBean = new BeanWithList("Main bean", new Double(10.0));
        listBean.addBean(simpleBean1);
        listBean.addBean(simpleBean2);
        listBean.addBean(simpleBean3);

        beans.put("bean", simpleBean1);
        beans.put("listBean", listBean);

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(iftagXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(iftagXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
//        assertEquals("Last Row Number is incorrect", 11, resultSheet.getLastRowNum());
        Map props = new HashMap();
        props.put( "${listBean.name}", listBean.getName());
        props.put( "${listBean.doubleValue}", listBean.getDoubleValue());
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows( sourceSheet, resultSheet, 4, 3, 1);
        checker.checkRows( sourceSheet, resultSheet, 4, 6, 1);
        checker.checkRows( sourceSheet, resultSheet, 4, 8, 1);
        checker.checkRows( sourceSheet, resultSheet, 8, 5, 1);
        checker.checkRows( sourceSheet, resultSheet, 8, 7, 1);
        checker.checkRows( sourceSheet, resultSheet, 8, 10, 1);
        props.clear();
        props.put("${sb.name}",names[0]);
        props.put("${sb.doubleValue}",doubleValues[0]);
        checker.checkRows( sourceSheet, resultSheet, 6, 4, 1);
        props.clear();
        props.put("${sb.name}",names[2]);
        props.put("${sb.doubleValue}",doubleValues[2]);
        checker.checkRows( sourceSheet, resultSheet, 6, 9, 1);

        is.close();
        saveWorkbook( resultWorkbook, iftagDestXLS);
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
        saveWorkbook( resultWorkbook, forifTag2DestXLS);
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

        saveWorkbook( resultWorkbook, forifTag3DestXLS);
    }
    
    public void testForIfTag3OutTag() throws IOException, ParsePropertyException {
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
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTag3OutTagXLS));
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

        saveWorkbook( resultWorkbook, forifTag3OutTagDestXLS);
    }

    public void testEmptyBeansExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();

        SimpleBean simpleBean1 = new SimpleBean(names[0].toString(), (Double) doubleValues[0], (Integer) intValues[0], (Date) dateValues[0]);

        BeanWithList listBean = new BeanWithList("Main bean", new Double(10.0));

        beans.put("bean", simpleBean1);
        beans.put("listBean", listBean);

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(emptyBeansXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(emptyBeansXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals("Last Row Number is incorrect", sourceSheet.getLastRowNum(), resultSheet.getLastRowNum());
        Map props = new HashMap();
        props.put("${listBean.name}", listBean.getName());
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);

        is.close();
        saveWorkbook( resultWorkbook, emptyBeansDestXLS );
    }


    public void testListOfStringsExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("employee", itEmployees.get(0));
        beans.put("employees", itEmployees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(employeeNotesXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(employeeNotesXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        checker.checkListCells( sourceSheet, 2, resultSheet, 2, (short)1, ((Employee)itEmployees.get(0)).getNotes().toArray() );

        is.close();
        saveWorkbook(resultWorkbook, employeeNotesDestXLS);

    }

    public void testExtendedEncodingExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        Employee emp = (Employee) itEmployees.get(0);
        emp.setName("Леонид");

        List notes = new ArrayList();
        notes.add("Запи�?ь 1");
        notes.add("Заметка 2");
        notes.add("Строка 3");
        emp.setNotes( notes );
        beans.put("employee", emp);
        beans.put("employees", itEmployees);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(employeeNotesXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(employeeNotesXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        Map props = new HashMap();
        CellsChecker checker = new CellsChecker(props);
        checker.checkListCells( sourceSheet, 2, resultSheet, 2, (short)1, emp.getNotes().toArray() );

        is.close();
        saveWorkbook( resultWorkbook, employeeNotesRusDestXLS);

    }

    public void testForIfTagOneRowExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRowXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRowXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);
        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);

        Map props = new HashMap();
        props.put( "${department.name}", "IT");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 1);

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
        saveWorkbook( resultWorkbook, forifTagOneRowDestXLS);
    }

    public void testDynamicColumns() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        List cols = new ArrayList();
        String[] colNames = new String[]{"Column 1", "Column 2", "Column 3"};
        for (int i = 0; i < colNames.length; i++) {
            String colName = colNames[i];
            cols.add( new Column(colName) );
        }
        beans.put( "cols", cols );
        List list = new ArrayList();
        list.add(new Item("A", new int[] { 1, 2, 3 }));
        list.add(new Item("B", new int[] { }));
        list.add(new Item("C", new int[] { 4, 5, 6 }));
        list.add(new Item("D", new int[] { }));
        beans.put("list", list);        

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(dynamicColumnsXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(dynamicColumnsXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);
        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);

        Map props = new HashMap();
        props.put( "${col.text}", colNames[0]);
        CellsChecker checker = new CellsChecker(props);
        checker.checkCells( sourceSheet, resultSheet, 0, (short)1, 0, (short)0, true);
        props.put( "${col.text}", colNames[1]);
        checker.checkCells( sourceSheet, resultSheet, 0, (short)1, 0, (short)1, true);
        props.put( "${col.text}", colNames[2]);
        checker.checkCells( sourceSheet, resultSheet, 0, (short)1, 0, (short)2, true);
        is.close();
        saveWorkbook( resultWorkbook, dynamicColumnsDestXLS);
    }

    public void testForIfTagOneRowExport2() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        List items = new ArrayList();
        items.add(new Item("Item 1"));
//        items.add(new Item("Item 2"));
        beans.put("items", items);

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRow2XLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        //todo: complete test
//        is = new BufferedInputStream(getClass().getResourceAsStream(forifTagOneRow2XLS));
//        POIFSFileSystem fs = new POIFSFileSystem(is);
//        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);
//        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
//        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);

//        is.close();
        saveWorkbook( resultWorkbook, forifTagOneRowDest2XLS);
    }

    public void testHiddenSheetsExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put("bean", simpleBean1);
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(hideSheetsXLS));
        XLSTransformer transformer = new XLSTransformer();
        transformer.setSpreadsheetsToRemove(new String[]{"Sheet 2", "Sheet 3"});
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        assertEquals("Number of sheets in result workbook is incorrect", 1, resultWorkbook.getNumberOfSheets() );
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(hideSheetsXLS));
        transformer.setSpreadsheetsToRemove(new String[]{"Sheet 2"});
        resultWorkbook = transformer.transformXLS(is, beans);
        assertEquals("Number of sheets in result workbook is incorrect", 2, resultWorkbook.getNumberOfSheets() );
        is.close();

        saveWorkbook(resultWorkbook, hideSheetsDestXLS);
    }

    public void testMultipleSheetList() throws IOException, ParsePropertyException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(multipleSheetListXLS));
        XLSTransformer transformer = new XLSTransformer();
        List sheetNames = new ArrayList();
//        sheetNames.add("New Sheet");
        for(int i = 0; i < departments.size(); i++){
            Department department = (Department) departments.get( i );
            sheetNames.add( department.getName() );
        }

        HSSFWorkbook resultWorkbook = transformer.transformMultipleSheetsList(is, departments, sheetNames, "department", new HashMap(), 0);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(multipleSheetListXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        assertEquals( "Number of result worksheets is incorrect ", sourceWorkbook.getNumberOfSheets() + departments.size() - 1, resultWorkbook.getNumberOfSheets());
//        for (int sheetNo = 0; sheetNo < resultWorkbook.getNumberOfSheets() && sheetNo < sheetNames.size(); sheetNo++) {
//             assertEquals( "Result worksheet name is incorrect", sheetNames.get(sheetNo), resultWorkbook.getSheetName(sheetNo));
//        }
// todo create all necessary checks
//        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
//        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
//
//        Map props = new HashMap();
//        props.put("${departments.name}//:4", "IT");
//        props.put("Department//departments", "Department");
//        CellsChecker checker = new CellsChecker(props);
//        checker.checkRows(sourceSheet, resultSheet, 0, 0, 3);
//        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 0, itEmployeeNames);
//        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 1, itPayments);
//        checker.checkListCells(sourceSheet, 3, resultSheet, 3, (short) 2, itBonuses);
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 3, (short)3, "B4*(1+C4)");
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 4, (short)3, "B5*(1+C5)");
//        checker.checkFormulaCell( sourceSheet, 3, resultSheet, 7, (short)3, "B8*(1+C8)");
        is.close();
        saveWorkbook(resultWorkbook, multipleSheetListDestXLS);
    }

    public void testMultiTab() throws IOException, ParsePropertyException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(multiTabXLS));
        XLSTransformer transformer = new XLSTransformer();
        List sheetNames = new ArrayList();
//        sheetNames.add("New Sheet");
        List maps = new ArrayList();
        for(int i = 0; i < departments.size(); i++){
            Map map = new HashMap();
            Department department = (Department) departments.get( i );
            map.put("department", department );
            sheetNames.add( department.getName() );
            map.put("name", "Number " + i);
            maps.add( map );
        }


        HSSFWorkbook resultWorkbook = transformer.transformMultipleSheetsList(is, maps, sheetNames, "map", new HashMap(), 0);
        is.close();
        saveWorkbook(resultWorkbook, multiTabDestXLS);
    }

    // todo complete this test
    public void atestMultipleSheetList2() throws IOException, ParsePropertyException {
        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(multipleSheetList2XLS));
        XLSTransformer transformer = new XLSTransformer();
        List sheetNames = new ArrayList();
        sheetNames.add("Sheet 1");
        for(int i = 0; i < departments.size(); i++){
            Department department = (Department) departments.get( i );
            sheetNames.add( department.getName() );
        }
        List templateSheetList = new ArrayList();
        templateSheetList.add("Template Sheet 1");
        templateSheetList.add("Template Sheet 2");
        List sheetNameList = new ArrayList();
        List beanParamList = new ArrayList();

        HSSFWorkbook resultWorkbook = transformer.transformMultipleSheetsList(is, departments, sheetNames, "department", new HashMap(), 0);
        transformer.transformXLS(is, templateSheetList, sheetNameList, beanParamList );
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(multipleSheetList2XLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        assertEquals( "Number of result worksheets is incorrect ", sourceWorkbook.getNumberOfSheets() + departments.size() - 1, resultWorkbook.getNumberOfSheets());
        for (int sheetNo = 0; sheetNo < resultWorkbook.getNumberOfSheets() && sheetNo < sheetNames.size(); sheetNo++) {
        }
        is.close();
        saveWorkbook(resultWorkbook, multipleSheetList2DestXLS);
    }

    public void testGroupTag() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(groupTagXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        // todo complete test checks
//        is = new BufferedInputStream(getClass().getResourceAsStream(groupTagXLS));
//        POIFSFileSystem fs = new POIFSFileSystem(is);
//        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

//        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
//        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);

        saveWorkbook( resultWorkbook, groupTagDestXLS );
    }

    public void testForIfTagMergeCellsExport() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forifTagMergeXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        saveWorkbook(resultWorkbook, forifTagMergeDestXLS);
    }

    public void testJEXLExpressions() throws IOException {
        Map beans = new HashMap();
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        beans.put("dateFormat", dateFormat);

        Map map = new HashMap();
        map.put("Name", "Leonid");
        map.put("Surname", "Vysochyn");
        map.put("employees", itDepartment.getStaff());

        beans.put("map", map);

        MyBean obj = new MyBean();

        Bean bean = new Bean();

        beans.put( "bean", bean );

        beans.put("obj", obj);
        beans.put("employees1", ((Department)departments.get(0)).getStaff());
        beans.put("employees2", new ArrayList());
        beans.put("employees3", ((Department)departments.get(1)).getStaff());
        beans.put("employees4", new ArrayList());
        beans.put("employees5", ((Department)departments.get(2)).getStaff());

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(jexlXLS));
        XLSTransformer transformer = new XLSTransformer();
        transformer.setJexlInnerCollectionsAccess( true );
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(jexlXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        Map props = new HashMap();
        props.put("${obj.name}", obj.getName() );
        props.put("${\"Hello, World\"}","Hello, World");
        props.put("${obj.flag == true}", Boolean.valueOf(obj.getFlag()));
        props.put("${obj.name == null}", Boolean.valueOf(obj.getName() == null));
//        props.put("${empty(obj.collection)}", Boolean.valueOf(obj.getCollection().isEmpty()));
//        props.put("${obj.collection.size()}", new Integer(((String)obj.getCollection().get(0)).length()));
        props.put("${obj.name.size()}", new Integer( obj.getName().length() ) );
        props.put("${!empty(obj.collection) && obj.id > 0}", Boolean.valueOf(!obj.getCollection().isEmpty() && obj.getId() > 0));
        props.put("${empty(obj.collection) || obj.id == 1}", Boolean.valueOf(obj.getCollection().isEmpty() && obj.getId() == 1));
        props.put("${not empty(obj.collection)}", Boolean.valueOf(!obj.getCollection().isEmpty()));
        props.put("${obj.id > 1}", Boolean.valueOf(obj.getId() > 1));
        props.put("${obj.id == 1}", Boolean.valueOf(obj.getId() == 1));
        props.put("${obj.id != 1}", Boolean.valueOf(obj.getId() != 1));
        props.put("${obj.id eq 1}", Boolean.valueOf(obj.getId() == 1));
        props.put("${obj.id % 2}", new Integer(obj.getId() % 2));
        props.put("${obj.myArray[0]} and ${obj.myArray[1]}", obj.getMyArray()[0] + " and " + obj.getMyArray()[1]);
        props.put("${dateFormat.format(obj.date)}", dateFormat.format( obj.getDate() ));
        props.put("${obj.printIt()}", obj.printIt() );
        props.put("${obj.getName()}", obj.getName());
        props.put("${obj.echo(\"Hello\")}", obj.echo("Hello"));

        CellsChecker checker = new CellsChecker(props);
        checker.checkSection( sourceSheet, resultSheet, 0, 0, (short)0, (short)1, 25, false, false);
        props.clear();
        props.put("${bean.collection.innerCollection.get(0)}", "1");
        checker.checkListCells( sourceSheet, 25, resultSheet, 25, (short)1,
                new String[]{((Bean.InnerBean)bean.getCollection().get(0)).getInnerCollection().get(0).toString(),
                              ((Bean.InnerBean)bean.getCollection().get(1)).getInnerCollection().get(0).toString(),
                                ((Bean.InnerBean)bean.getCollection().get(2)).getInnerCollection().get(0).toString()});

        saveWorkbook( resultWorkbook, jexlDestXLS);

    }

    public void testForGroupBy() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        List deps = new ArrayList( departments );
        // adding department with null values to check grouping with null values
        deps.add(mgrDepartment);
        beans.put( "departments", deps );

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        is.close();
        saveWorkbook(resultWorkbook, forGroupByDestXLS);
        // testing empty collection used with jx:forEach grouping 
//        Collection emptyDepartments = new ArrayList();
        ((Department)departments.get(0)).getStaff().clear();
//        beans.clear();
//        beans.put( "departments", emptyDepartments );
        is = new BufferedInputStream(getClass().getResourceAsStream(forGroupByXLS));
        resultWorkbook = transformer.transformXLS(is, beans);
        saveWorkbook( resultWorkbook, forGroupByDestXLS );
        is.close();

    }

    public void testPoiObjectsExpose() throws IOException, ParsePropertyException {
        Map beans = new HashMap();
        beans.put( "departments", departments );
        beans.put("itDepartment", itDepartment);

        List employees = itDepartment.getStaff();
        ((Employee)employees.get(0)).setComment("");
        for (int i = 1; i < employees.size(); i++) {
            Employee employee = (Employee) employees.get(i);
            String comment = "";
            for( int j =0; j <= i; j++ ){
                comment += "Employee Comment Line " + j + " ..\r\n";
            }
            employee.setComment( comment );
        }
        beans.put("employees", employees);
        beans.put("lineSize", new Integer(0));
        beans.put("row", new Integer(3));

        InputStream is = new BufferedInputStream(getClass().getResourceAsStream(poiobjectsXLS));
        XLSTransformer transformer = new XLSTransformer();
        HSSFWorkbook resultWorkbook = transformer.transformXLS(is, beans);
        is.close();
        is = new BufferedInputStream(getClass().getResourceAsStream(poiobjectsXLS));
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fs);

        HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
        HSSFSheet resultSheet = resultWorkbook.getSheetAt(0);
        assertEquals("First Row Numbers differ in source and result sheets", sourceSheet.getFirstRowNum(), resultSheet.getFirstRowNum());
        assertEquals(resultSheet.getHeader().getLeft(), "Test Left Header");
        assertEquals(resultSheet.getHeader().getCenter(), itDepartment.getName());
        assertEquals(resultSheet.getHeader().getRight(), "Test Right Header");
        assertEquals(resultSheet.getFooter().getRight(), "Test Right Footer");
        assertEquals(resultSheet.getFooter().getCenter(), "Test Center Footer");
        assertEquals( resultWorkbook.getSheetName(2), itDepartment.getName());
        Map props = new HashMap();
        props.put("${department.name}", "IT");
        CellsChecker checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 0, 3);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 0, itEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 1, itPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 3, (short) 2, itBonuses);
        props.clear();
        props.put("${department.name}", "HR");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 9, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 0, hrEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 1, hrPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 12, (short) 2, hrBonuses);
        props.clear();
        props.put("${department.name}", "BA");
        checker = new CellsChecker(props);
        checker.checkRows(sourceSheet, resultSheet, 1, 17, 3);
        props.clear();
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 0, baEmployeeNames);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 1, baPayments);
        checker.checkListCells(sourceSheet, 5, resultSheet, 20, (short) 2, baBonuses);
        is.close();
        saveWorkbook( resultWorkbook, poiobjectsDestXLS);
    }


    private void saveWorkbook(HSSFWorkbook resultWorkbook, String fileName) throws IOException {
        String saveResultsProp = System.getProperty("saveResults");
        if( "true".equalsIgnoreCase(saveResultsProp) ){
            if( log.isInfoEnabled() ){
                log.info("Saving " + fileName);
            }
            OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
            resultWorkbook.write(os);
            os.flush();
            os.close();
            log.info("Output Excel saved to " + fileName);
        }
    }


}
