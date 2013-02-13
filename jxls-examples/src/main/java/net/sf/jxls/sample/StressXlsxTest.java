package net.sf.jxls.sample;

import net.sf.jxls.sample.model.Department;
import net.sf.jxls.sample.model.Employee;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class StressXlsxTest {

    private static String templateFileDir;
    private static String destFileDir;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileDir = args[0];
            destFileDir = args[1];
        }
        StressXlsxTest test = new StressXlsxTest();
        test.testStress1();
        test.testStress2();
    }

    public void testStress1() throws InvalidFormatException, IOException {
        Map beans = new HashMap();
        final int employeeCount = 30000;
        List<Employee> employees = Employee.generate(employeeCount);
        beans.put("employees", employees);
        InputStream is = new BufferedInputStream(new FileInputStream(templateFileDir + "stress1.xlsx"));
        XLSTransformer transformer = new XLSTransformer();
        long startTime = System.nanoTime();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        long endTime = System.nanoTime();
        is.close();
        saveWorkbook(resultWorkbook, destFileDir + "stress1_output.xlsx");
        System.out.println("Stress1 XLSX time (s): " + (endTime - startTime)/1000000000);
    }

    public void testStress2() throws InvalidFormatException, IOException {
        Map beans = new HashMap();
        final int employeeCount = 500;
        final int depCount = 100;
        List<Department> departments = Department.generate(depCount, employeeCount);
        beans.put("departments", departments);
        InputStream is = new BufferedInputStream(new FileInputStream(templateFileDir + "stress2.xlsx"));
        XLSTransformer transformer = new XLSTransformer();
        long startTime = System.nanoTime();
        Workbook resultWorkbook = transformer.transformXLS(is, beans);
        long endTime = System.nanoTime();
        is.close();
        saveWorkbook(resultWorkbook, destFileDir + "stress2_output.xlsx");
        System.out.println("Stress2 XLSX time (s): " + (endTime - startTime)/1000000000);
    }

    private void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
        OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
        resultWorkbook.write(os);
        os.flush();
        os.close();
    }
}
