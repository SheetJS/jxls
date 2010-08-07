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

/**
 * @author Leonid Vysochyn
 */
public class XLSXReaderTest extends TestCase {
    public static final String dataXLS = "/templates/departmentdata.xlsx";

    public void testRead() throws IOException, InvalidFormatException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(dataXLS));

        Department itDepartment = new Department();
        Department hrDepartment = new Department();
        Map beans = new HashMap();
        beans.put("itDepartment", itDepartment);
        beans.put("hrDepartment", hrDepartment);
        // Create Sheet1 Reader
        List chiefMappings = new ArrayList();
        chiefMappings.add( new BeanCellMapping(0, (short) 1, "itDepartment", "name") );
        chiefMappings.add( new BeanCellMapping(3, (short) 0, "itDepartment", "chief.name") );
        chiefMappings.add( new BeanCellMapping(3, (short) 1, "itDepartment", "chief.age") );
        chiefMappings.add( new BeanCellMapping(3, (short) 3, "itDepartment", "chief.payment") );
        chiefMappings.add( new BeanCellMapping("E4", "itDepartment", "chief.bonus") );
        XLSBlockReader chiefReader = new SimpleBlockReaderImpl(0, 6, chiefMappings);
        List employeeMappings = new ArrayList();
        employeeMappings.add( new BeanCellMapping(7, (short) 0, "employee", "name") );
        employeeMappings.add( new BeanCellMapping(7, (short) 1, "employee", "age") );
        employeeMappings.add( new BeanCellMapping(7, (short) 3, "employee", "payment") );
        employeeMappings.add( new BeanCellMapping(7, (short) 4, "employee", "bonus") );
        XLSBlockReader employeeReader = new SimpleBlockReaderImpl(7, 7, employeeMappings);
        XLSLoopBlockReader employeesReader = new XLSForEachBlockReaderImpl(7, 7, "itDepartment.staff", "employee", Employee.class);
        employeesReader.addBlockReader( employeeReader );
        SectionCheck loopBreakCheck = getLoopBreakCheck();
        employeesReader.setLoopBreakCondition( loopBreakCheck );
        XLSSheetReader sheet1Reader = new XLSSheetReaderImpl();
        sheet1Reader.addBlockReader( chiefReader );
        sheet1Reader.addBlockReader( employeesReader );
        // Create Sheet2 Reader
        XLSSheetReader sheet2Reader = new XLSSheetReaderImpl();
        employeeMappings = new ArrayList();
        employeeMappings.add( new BeanCellMapping(2, (short) 0, "employee", "name") );
        employeeMappings.add( new BeanCellMapping(2, (short) 1, "employee", "age") );
        employeeMappings.add( new BeanCellMapping(2, (short) 2, "employee", "payment") );
        employeeMappings.add( new BeanCellMapping(2, (short) 3, "employee", "bonus") );
        XLSBlockReader sheet2EmployeeReader = new SimpleBlockReaderImpl(2, 2, employeeMappings);
        XLSLoopBlockReader sheet2EmployeesReader = new XLSForEachBlockReaderImpl(2, 2, "hrDepartment.staff", "employee", Employee.class);
        sheet2EmployeesReader.addBlockReader( sheet2EmployeeReader );
        sheet2EmployeesReader.setLoopBreakCondition( getLoopBreakCheck() );
        chiefMappings = new ArrayList();
        chiefMappings.add( new BeanCellMapping(7, (short)0, "hrDepartment", "chief.name"));
        chiefMappings.add( new BeanCellMapping(7, (short)1, "hrDepartment", "chief.age"));
        chiefMappings.add( new BeanCellMapping(7, (short)2, "hrDepartment", "chief.payment"));
        chiefMappings.add( new BeanCellMapping(7, (short)3, "hrDepartment", "chief.bonus"));
        XLSBlockReader hrChiefReader = new SimpleBlockReaderImpl(3, 7, chiefMappings);
        sheet2Reader.addBlockReader( new SimpleBlockReaderImpl(0, 1, new ArrayList()));
        sheet2Reader.addBlockReader( sheet2EmployeesReader );
        sheet2Reader.addBlockReader( hrChiefReader );
        // create main reader
        XLSReader mainReader = new XLSReaderImpl();
        mainReader.addSheetReader("Sheet1", sheet1Reader);
        mainReader.addSheetReader("Sheet2", sheet2Reader);
        mainReader.read( inputXLS, beans);
        inputXLS.close();
        // check sheet1 data
        assertEquals( "IT", itDepartment.getName() );
        assertEquals( "Maxim", itDepartment.getChief().getName() );
        assertEquals( new Integer(30), itDepartment.getChief().getAge() );
        assertEquals( new Double( 3000.0), itDepartment.getChief().getPayment() );
        assertEquals( new Double(0.25), itDepartment.getChief().getBonus() );
        assertEquals( 4, itDepartment.getStaff().size() );
        Employee employee = (Employee) itDepartment.getStaff().get(0);
        checkEmployee( employee, "Oleg", new Integer(32), new Double(2000.0), new Double(0.20) );
        employee = (Employee) itDepartment.getStaff().get(1);
        checkEmployee( employee, "Yuri", new Integer(29), new Double(1800.0), new Double(0.15) );
        employee = (Employee) itDepartment.getStaff().get(2);
        checkEmployee( employee, "Leonid", new Integer(30), new Double(1700.0), new Double(0.20) );
        employee = (Employee) itDepartment.getStaff().get(3);
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


    private SectionCheck getLoopBreakCheck() {
        OffsetRowCheck rowCheck = new OffsetRowCheckImpl( 0 );
        rowCheck.addCellCheck( new OffsetCellCheckImpl((short) 0, "Employee Payment Totals:") );
        SectionCheck sectionCheck = new SimpleSectionCheck();
        sectionCheck.addRowCheck( rowCheck );
        return sectionCheck;
    }


}
