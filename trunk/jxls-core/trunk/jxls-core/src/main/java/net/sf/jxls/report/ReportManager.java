package net.sf.jxls.report;

import java.util.List;
import java.sql.SQLException;

/**
 * @author Leonid Vysochyn
 */
public interface ReportManager {
    public List exec(String sql) throws SQLException;
}
