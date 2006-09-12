package net.sf.jxls.sample;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;
import net.sf.jxls.report.ReportManager;
import net.sf.jxls.report.ReportManagerImpl;

import java.sql.*;
import java.util.Map;
import java.util.HashMap;

/**
 * This sample demonstrates reporting capabilities of jXLS
 * @author Leonid Vysochyn
 */
public class ReportingSample {
    private static String templateFileName = "examples/templates/report.xls";
    private static String destFileName = "build/report_output.xls";

    public static void main(String[] args) throws Exception, ParsePropertyException, ClassNotFoundException, SQLException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        DatabaseHelper dbHelper = new DatabaseHelper();
        Class.forName("org.hsqldb.jdbcDriver");
        Connection conn = DriverManager.getConnection("jdbc:hsqldb:mem:jxls", "sa", "");
        dbHelper.initDatabase( conn );
        Map beans = new HashMap();
        ReportManager reportManager = new ReportManagerImpl( conn, beans );
        beans.put("rm", reportManager);
        beans.put("minDate", "1979-01-01");
        XLSTransformer transformer = new XLSTransformer();
        transformer.transformXLS(templateFileName, beans, destFileName);
    }

}
