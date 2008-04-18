package net.sf.jxls.sample;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.transformer.XLSTransformer;

import org.apache.commons.beanutils.RowSetDynaClass;

/**
 * @author Leonid Vysochyn
 */
public class RowSetExport {

    private static String templateFileName = "examples/templates/employees.xls";
    private static String destFileName = "build/employees_output.xls";

    public static void main(String[] args) throws Exception,  ClassNotFoundException, SQLException {
        if (args.length >= 2) {
            templateFileName = args[0];
            destFileName = args[1];
        }
        DatabaseHelper dbHelper = new DatabaseHelper();
        Class.forName("org.hsqldb.jdbcDriver");
        Connection con = DriverManager.getConnection("jdbc:hsqldb:mem:jxls", "sa", "");
        dbHelper.initDatabase( con );
        // get result set
        Statement stmt = con.createStatement();
        String query = "SELECT name, age, payment, bonus, birthDate FROM employee";
        ResultSet rs = stmt.executeQuery(query);
        // second parameter to RowSetDynaClass constructor indicates should the properties be lowercased
        RowSetDynaClass rsdc = new RowSetDynaClass(rs, true);
        Map beans = new HashMap();
        beans.put("employee", rsdc.getRows());
        XLSTransformer transformer = new XLSTransformer();
        transformer.transformXLS(templateFileName, beans, destFileName);
        stmt.close();
        rs.close();
        con.close();
    }
}
