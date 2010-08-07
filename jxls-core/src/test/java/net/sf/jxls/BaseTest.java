package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.bean.Department;
import net.sf.jxls.bean.Employee;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
* @author Leonid Vysochyn
*/
public abstract class BaseTest extends TestCase {
    protected final Log log = LogFactory.getLog(getClass());

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
    
    protected void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
            OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
            resultWorkbook.write(os);
            os.flush();
            os.close();
    }

}
