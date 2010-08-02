package net.sf.jxls.sample;

import java.io.IOException;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.sample.model.Employee;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * @author Leonid Vysochyn
 */
public class SingleListExport {
    private static String templateFileName = "examples/templates/employees.xls";
    private static String destFileName = "build/employees_output.xls";

    public static void main(String[] args) throws IOException, ParsePropertyException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        Collection staff = new HashSet();
        staff.add(new Employee("Derek", 35, 3000, 0.30));
        staff.add(new Employee("Elsa", 28, 1500, 0.15));
        staff.add(new Employee("Oleg", 32, 2300, 0.25));
        staff.add(new Employee("Neil", 34, 2500, 0.00));
        staff.add(new Employee("Maria", 34, 1700, 0.15));
        staff.add(new Employee("John", 35, 2800, 0.20));
        Map beans = new HashMap();
        beans.put("employee", staff);
        Configuration config = new Configuration();
//        config.setUTF16( true );
        XLSTransformer transformer = new XLSTransformer( config );
        transformer.groupCollection( "employee.name");
        transformer.transformXLS(templateFileName, beans, destFileName);
    }

}
