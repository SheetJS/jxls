package net.sf.jxls.sample;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.report.ResultSetCollection;
import net.sf.jxls.transformer.XLSTransformer;

/**
 * @author Leonid Vysochyn
 */
public class ResultSetExport {
    private static String templateFileName = "examples/templates/employees.xls";
    private static String destFileName = "build/resultset_output.xls";

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
        // let's get number of records to be able to implement size() method in ResultSetCollection class
        String countQuery = "SELECT COUNT(*) FROM employee";
        ResultSet rs = stmt.executeQuery( countQuery );
        int count = 0;
        if( rs.next() ){
            count = rs.getInt( 1 );
        }

        String query = "SELECT name, age, payment, bonus, birthDate FROM employee";
        rs = stmt.executeQuery(query);
        Map beans = new HashMap();
        ResultSetCollection rsc = new ResultSetCollection(rs, count, true);
        beans.put("employee", rsc);
        XLSTransformer transformer = new XLSTransformer();
        transformer.transformXLS(templateFileName, beans, destFileName);
        stmt.close();
        rs.close();
        con.close();
    }
}
