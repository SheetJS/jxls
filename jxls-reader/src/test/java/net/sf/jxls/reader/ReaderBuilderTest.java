package net.sf.jxls.reader;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import junit.framework.TestCase;
import net.sf.jxls.reader.sample.Department;
import net.sf.jxls.reader.sample.Employee;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;

/**
 * @author Leonid Vysochyn
 */
public class ReaderBuilderTest extends TestCase {
    public static final String dataXLS = "/templates/departmentData.xls";
    public static final String xmlConfig = "/xml/departments.xml";

    public void testBuildFromXML() throws IOException, SAXException {
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig));
        XLSReader reader = ReaderBuilder.buildFromXML( inputXML );
        assertNotNull( reader );
        assertEquals( 3, reader.getSheetReaders().size() );
        XLSSheetReader sheetReader1 = (XLSSheetReader) reader.getSheetReaders().get("Sheet1");
        assertEquals( 2, sheetReader1.getBlockReaders().size() );
        XLSBlockReader blockReader = (XLSBlockReader) sheetReader1.getBlockReaders().get(0);
        assertTrue( blockReader instanceof SimpleBlockReader);
        SimpleBlockReader blockReader1 = (SimpleBlockReader) blockReader;
        assertEquals( 5, blockReader1.getMappings().size() );
        BeanCellMapping mapping1 = (BeanCellMapping) blockReader1.getMappings().get(0);
        assertEquals( "department.name", mapping1.getFullPropertyName() );
        assertEquals( "B1", mapping1.getCellName() );
        BeanCellMapping mapping2 = (BeanCellMapping) blockReader1.getMappings().get(1);
        assertEquals( "department.chief.name", mapping2.getFullPropertyName() );
        assertEquals( "A4", mapping2.getCellName() );
        XLSSheetReader sheetReader2 = (XLSSheetReader) reader.getSheetReaders().get("Sheet2");
        assertEquals( 3, sheetReader2.getBlockReaders().size() );
        XLSSheetReader sheetReader3 = (XLSSheetReader) reader.getSheetReaders().get("Sheet3");
        assertEquals( 1, sheetReader3.getBlockReaders().size() );
    }

    public void testBuildFromXML2() throws IOException, SAXException, InvalidFormatException {
        InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(xmlConfig));
        XLSReader mainReader = ReaderBuilder.buildFromXML( inputXML );
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(dataXLS));
        Department department = new Department();
        Department hrDepartment = new Department();
        List departments = new ArrayList();
        Map beans = new HashMap();
        beans.put("department", department);
        beans.put("hrDepartment", hrDepartment);
        beans.put("departments", departments);
        mainReader.read( inputXLS, beans);
        inputXLS.close();
        // check sheet1 data
        assertEquals( "IT", department.getName() );
        assertEquals( "Maxim", department.getChief().getName() );
        assertEquals( new Integer(30), department.getChief().getAge() );
        assertEquals( new Double( 3000.0), department.getChief().getPayment() );
        assertEquals( new Double(0.25), department.getChief().getBonus() );
        assertEquals( 4, department.getStaff().size() );
        Employee employee = (Employee) department.getStaff().get(0);
        checkEmployee( employee, "Oleg", new Integer(32), new Double(2000.0), new Double(0.20) );
        employee = (Employee) department.getStaff().get(1);
        checkEmployee( employee, "Yuri", new Integer(29), new Double(1800.0), new Double(0.15) );
        employee = (Employee) department.getStaff().get(2);
        checkEmployee( employee, "Leonid", new Integer(30), new Double(1700.0), new Double(0.20) );
        employee = (Employee) department.getStaff().get(3);
        checkEmployee( employee, "Alex", new Integer(28), new Double(1600.0), new Double(0.20) );
        // check sheet2 data
        checkEmployee( hrDepartment.getChief(), "Betsy", new Integer(37), new Double(2200.0), new Double(0.3) );
        assertEquals(4, hrDepartment.getStaff().size() );
        employee = (Employee) hrDepartment.getStaff().get(0);
        checkEmployee( employee, "Olga", new Integer(26), new Double(1400.0), new Double(0.20) );
        employee = (Employee) hrDepartment.getStaff().get(1);
        checkEmployee( employee, "Helen", new Integer(30), new Double(2100.0), new Double(0.10) );
        employee = (Employee) hrDepartment.getStaff().get(2);
        checkEmployee( employee, "Keith", new Integer(24), new Double(1800.0), new Double(0.15) );
        employee = (Employee) hrDepartment.getStaff().get(3);
        checkEmployee( employee, "Cat", new Integer(34), new Double(1900.0), new Double(0.15) );
    }

    private void checkEmployee(Employee employee, String name, Integer age, Double payment, Double bonus){
        assertNotNull( employee );
        assertEquals( name, employee.getName() );
        assertEquals( age, employee.getAge() );
        assertEquals( payment, employee.getPayment() );
        assertEquals( bonus, employee.getBonus() );
    }

    public void atestBuildWorkbookReader(){
        ReaderBuilder builder = new ReaderBuilder();
        builder.addSheetReader("Sheet 1").addSheetReader("Sheet 2");
        XLSReader reader = builder.getReader();
        assertEquals( "Incorrect number of sheet readers created", reader.getSheetReaders().size(), 2);
    }

    public void atestBuildWorkbookSheetBlockReader(){
        ReaderBuilder builder = new ReaderBuilder();
        builder.addSheetReader("Sheet 1");
        builder.addSimpleBlockReader( 1, 2 );
        builder.addSimpleBlockReader( 4, 5 );
        XLSReader reader = builder.getReader();
        XLSSheetReader sheetReader = (XLSSheetReader) reader.getSheetReaders().get("Sheet 1");
        assertEquals("Incorrect number of block readers", sheetReader.getBlockReaders().size(), 2);
        XLSBlockReader blockReader1 = (XLSBlockReader) sheetReader.getBlockReaders().get(0);
        assertEquals("BlockReader start row is incorrect", blockReader1.getStartRow(), 1);
        assertEquals("BlockReader end row is incorrect", blockReader1.getEndRow(), 2);
    }

    public void atestBuildWorkbookSheetBlockReaderWithMappings(){
        ReaderBuilder builder = new ReaderBuilder();
        builder.addSheetReader("Sheet 1");
        builder.addSimpleBlockReader( 0, 6 );
        builder.addMapping("B1", "department.name");
        builder.addMapping("A4", "department.chief.name");
        builder.addMapping("D4", "department.chief.payment");
        XLSReader reader = builder.getReader();
        XLSSheetReader sheetReader = (XLSSheetReader) reader.getSheetReaders().get("Sheet 1");
        XLSBlockReader blockReader = (XLSBlockReader) sheetReader.getBlockReaders().get(0);
//        assertEquals("Number of mappings is incorrect", blockReader.getMappings().size(), 3 );
//        BeanCellMapping mapping = (BeanCellMapping) blockReader.getMappings().get(0);
//        assertEquals("Incorrect cell name", mapping.getCellName(), "B1");
//        assertEquals("Incorrect property name", mapping.getFullPropertyName(), "department.name");
//        mapping = (BeanCellMapping) blockReader.getMappings().get(2);
//        assertEquals("Incorrect cell name", mapping.getCellName(), "D4");
//        assertEquals("Incorrect property name", mapping.getFullPropertyName(), "department.chief.payment");
    }

    public void atestBuildWorkbookSheetLoopReader(){
        ReaderBuilder builder = new ReaderBuilder();
        builder.addSheetReader("Sheet 1");
        builder.addLoopBlockReader(7, 8, "department.staff", "employee", Employee.class);
        builder.addSimpleBlockReader( 7, 10 );
        builder.addSimpleBlockReader( 11, 15 );
        XLSReader reader = builder.getReader();
        XLSSheetReader sheetReader = (XLSSheetReader) reader.getSheetReaders().get("Sheet 1");
        XLSLoopBlockReader loopBlockReader = (XLSLoopBlockReader) sheetReader.getBlockReaders().get(0);
        assertTrue( loopBlockReader instanceof XLSForEachBlockReaderImpl );

        assertEquals("Number of block readers is incorrect", loopBlockReader.getBlockReaders().size(), 2 );
        XLSBlockReader blockReader1 = (XLSBlockReader) loopBlockReader.getBlockReaders().get(0);
        assertEquals("BlockReader start row is incorrect", blockReader1.getStartRow(), 7);
        assertEquals("BlockReader end row is incorrect", blockReader1.getEndRow(), 10);
        XLSBlockReader blockReader2 = (XLSBlockReader) loopBlockReader.getBlockReaders().get(1);
        assertEquals("BlockReader start row is incorrect", blockReader2.getStartRow(), 11);
        assertEquals("BlockReader end row is incorrect", blockReader2.getEndRow(), 15);

    }

    public void atestBuildFunctionalWorkbookReader(){
        ReaderBuilder builder = new ReaderBuilder();
        builder.addSheetReader("Sheet 1");
        builder.addSimpleBlockReader( 0, 6 );
        builder.addMapping("B1", "department.name");
        builder.addMapping("A4", "department.chief.name");

        builder.addSheetReader("Sheet 2");
    }
}
