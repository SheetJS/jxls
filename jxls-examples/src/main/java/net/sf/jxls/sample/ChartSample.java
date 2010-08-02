package net.sf.jxls.sample;


import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.sample.model.Employee;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * @author Leonid Vysochyn
 */
public class ChartSample {
    private static String templateFileName = "examples/templates/chart.xls";
    private static String destFileName = "build/chart_output.xls";

    public static void main(String[] args) throws IOException, ParsePropertyException, InvalidFormatException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        List staff = new ArrayList();
        staff.add(new Employee("Derek", 35, 3000, 0.30));
        staff.add(new Employee("Elsa", 28, 1500, 0.15));
        staff.add(new Employee("Oleg", 32, 2300, 0.25));
        staff.add(new Employee("Neil", 34, 2500, 0.00));
        staff.add(new Employee("Maria", 34, 1700, 0.15));
        staff.add(new Employee("John", 35, 2800, 0.20));
        staff.add(new Employee("Leonid", 29, 1700, 0.20));
        Map beans = new HashMap();
        beans.put("employee", staff);
        XLSTransformer transformer = new XLSTransformer();
        transformer.markAsFixedSizeCollection("employee");
        transformer.transformXLS(templateFileName, beans, destFileName);
    }
}
