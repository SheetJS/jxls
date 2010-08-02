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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author Leonid Vysochyn
 */
public class XLSSheetReaderTest extends TestCase {
    public static final String dataXLS = "/templates/departmentData.xls";

    public void testRead() throws IOException, InvalidFormatException {
        InputStream inputXLS = new BufferedInputStream(getClass().getResourceAsStream(dataXLS));
        Workbook hssfInputWorkbook = WorkbookFactory.create(inputXLS);
        Sheet sheet = hssfInputWorkbook.getSheetAt( 0 );

        Department department = new Department();
        Employee chief = new Employee();
        Map beans = new HashMap();
        beans.put("department", department);
        beans.put("chief", chief);
        List chiefMappings = new ArrayList();
        chiefMappings.add( new BeanCellMapping(0, (short) 1, "department", "name") );
        chiefMappings.add( new BeanCellMapping(3, (short) 0, "chief", "name") );
        chiefMappings.add( new BeanCellMapping(3, (short) 1, "chief", "age") );
        chiefMappings.add( new BeanCellMapping(3, (short) 3, "chief", "payment") );
        chiefMappings.add( new BeanCellMapping(3, (short) 4, "chief", "bonus") );
        XLSBlockReader reader1 = new SimpleBlockReaderImpl(0, 6, chiefMappings);
        
        List employeeMappings = new ArrayList();
        employeeMappings.add( new BeanCellMapping(7, (short) 0, "employee", "name") );
        employeeMappings.add( new BeanCellMapping(7, (short) 1, "employee", "age") );
        employeeMappings.add( new BeanCellMapping(7, (short) 3, "employee", "payment") );
        employeeMappings.add( new BeanCellMapping(7, (short) 4, "employee", "bonus") );

        XLSBlockReader reader = new SimpleBlockReaderImpl(7, 7, employeeMappings);
        XLSLoopBlockReader forEachReader = new XLSForEachBlockReaderImpl(7, 7, "department.staff", "employee", Employee.class);
        forEachReader.addBlockReader( reader );
        SectionCheck loopBreakCheck = getLoopBreakCheck();
        forEachReader.setLoopBreakCondition( loopBreakCheck );

        XLSSheetReader sheetReader = new XLSSheetReaderImpl();
        sheetReader.addBlockReader( reader1 );
        sheetReader.addBlockReader( forEachReader );
        sheetReader.read( sheet, beans );

        assertEquals( "IT", department.getName() );
        assertEquals( "Maxim", chief.getName() );
        assertEquals( new Integer(30), chief.getAge() );
        assertEquals( new Double( 3000.0), chief.getPayment() );
        assertEquals( new Double(0.25), chief.getBonus() );

        assertEquals( 4, department.getStaff().size() );
        Employee employee = (Employee) department.getStaff().get(0);
        checkEmployee( employee, "Oleg", new Integer(32), new Double(2000.0), new Double(0.20) );
        employee = (Employee) department.getStaff().get(1);
        checkEmployee( employee, "Yuri", new Integer(29), new Double(1800.0), new Double(0.15) );
        employee = (Employee) department.getStaff().get(2);
        checkEmployee( employee, "Leonid", new Integer(30), new Double(1700.0), new Double(0.20) );
        employee = (Employee) department.getStaff().get(3);
        checkEmployee( employee, "Alex", new Integer(28), new Double(1600.0), new Double(0.20) );
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
