package net.sf.jxls.sample;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.sample.model.Department;
import net.sf.jxls.sample.model.Employee;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * @author Leonid Vysochyn
 */
public class AdjacentListsSample {
    private static String templateFileName = "/classes/templates/adjacentlists.xls";
    private static String destFileName = "build/adjacentlists_output.xls";

    public static void main(String[] args) throws IOException, ParsePropertyException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        Department depIT = new Department("IT");

        Employee chief = new Employee("Derek", 35, 3000, 0.30);
        depIT.setChief(chief);
        Employee elsa = new Employee("Elsa", 28, 1500, 0.15);
        depIT.addEmployee(elsa);
        Employee oleg = new Employee("Oleg", 32, 2300, 0.25);
        depIT.addEmployee(oleg);
        Employee neil = new Employee("Neil", 34, 2500, 0.00);
        depIT.addEmployee(neil);
        Employee maria = new Employee("Maria", 34, 1700, 0.15);
        depIT.addEmployee(maria);
        Employee john = new Employee("John", 35, 2800, 0.20);
        depIT.addEmployee(john);

        Department depHR = new Department("HR");

        Employee natali = new Employee("Natali", 25, 1200, 0.1);
        depHR.addEmployee( natali );
        Employee helen = new Employee("Helen", 27, 1100, 0.20);
        depHR.addEmployee(helen);
        Employee olga = new Employee("Olga", 24, 1150, 0.00);
        depHR.addEmployee(olga);

        Map beans = new HashMap();
        beans.put("depIT", depIT);
        beans.put("depHR", depHR);

        XLSTransformer transformer = new XLSTransformer();
        transformer.transformXLS(templateFileName, beans, destFileName);
    }
}
