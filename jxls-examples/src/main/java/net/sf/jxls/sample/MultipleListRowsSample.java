package net.sf.jxls.sample;

import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
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
public class MultipleListRowsSample {
    private static String templateFileName = "examples/templates/multiplelistrows.xls";
    private static String destFileName = "build/multiplelistrows_output.xls";

    public static void main(String[] args) throws IOException, ParsePropertyException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        Department department = new Department("IT");
        Calendar calendar = Calendar.getInstance();
        calendar.set(1970, 12, 2);
        Date d1 = calendar.getTime();
        calendar.set(1980, 2, 15);
        Date d2 = calendar.getTime();
        calendar.set(1976, 7, 20);
        Date d3 = calendar.getTime();
        calendar.set(1968, 5, 6);
        Date d4 = calendar.getTime();
        calendar.set(1978, 8, 17);
        Date d5 = calendar.getTime();
        Employee chief = new Employee("Derek", 35, 3000, 0.30, d1);
        department.setChief(chief);
        Employee elsa = new Employee("Elsa", 28, 1500, 0.15, d2);
        department.addEmployee(elsa);
        Employee oleg = new Employee("Oleg", 32, 2300, 0.25, d3);
        department.addEmployee(oleg);
        Employee neil = new Employee("Neil", 34, 2500, 0.00, d4);
        department.addEmployee(neil);
        Employee maria = new Employee("Maria", 34, 1700, 0.15, d5);
        department.addEmployee(maria);
        Employee john = new Employee("John", 35, 2800, 0.20, d2);
        department.addEmployee(john);
        maria.setSuperior(oleg);
        oleg.setSuperior(john);
        neil.setSuperior(john);
        Map beans = new HashMap();
        beans.put("department", department);
        XLSTransformer transformer = new XLSTransformer();
        //transformer.registerCellProcessor( new RedCellProcessor("department.staff") );
        //transformer.registerCellProcessor( new MultPropertyCellProcessor("department.staff", adjProperties ));
        //transformer.registerCellProcessor( new MultPropertyCellProcessor("department.chief", adjProperties));
//        transformer.groupCollection("department.staff");
        transformer.transformXLS(templateFileName, beans, destFileName);
    }

}
