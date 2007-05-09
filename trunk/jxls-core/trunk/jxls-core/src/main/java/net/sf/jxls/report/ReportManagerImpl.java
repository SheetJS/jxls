package net.sf.jxls.report;

import org.apache.commons.beanutils.RowSetDynaClass;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.util.List;
import java.util.Map;
import java.sql.Connection;
import java.sql.Statement;
import java.sql.SQLException;
import java.sql.ResultSet;

/**
 * @author Leonid Vysochyn
 */
public class ReportManagerImpl implements ReportManager {
    protected final Log log = LogFactory.getLog(getClass());
    Connection connection;
    Map beans;

    public ReportManagerImpl(Connection connection, Map beans) {
        this.connection = connection;
        this.beans = beans;
    }

    public ReportManagerImpl(Connection connection) {
        this.connection = connection;
    }

    public Connection getConnection() {
        return connection;
    }

    public void setConnection(Connection connection) {
        this.connection = connection;
    }

    public List exec(String sql) throws SQLException {
        Statement stmt = connection.createStatement();
        sql = sql.replaceAll("&apos;", "'");
        ResultSet rs = stmt.executeQuery(sql);
        // second parameter to RowSetDynaClass constructor indicates that properties should not be lowercased
        RowSetDynaClass rsdc = new RowSetDynaClass(rs, true);
        stmt.close();
        rs.close();
        return rsdc.getRows();
    }


}
