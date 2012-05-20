package net.sf.jxls;

import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class RepeatedFormulaTest {
    public static final String templateFileName = "/templates/repeatedFormula.xls";
    public static final String outputFileName = "target/repeatedFormula_output.xls";

    public static class FooSampleReportBean {
        public List<FooSampleReportRow> rowSampleSet1 = new ArrayList<FooSampleReportRow>();
    }

    public static class FooSampleReportRow {
        public int f0Int;

        public FooSampleReportRow( int f0 ) {
            f0Int = f0;
        }

        public int getF0Int() {
            return f0Int;
        }
    }

    private FooSampleReportBean report;

    @Before
    public void setUp() throws Exception {
        createSample();
    }

    private void createSample() {
        report = new FooSampleReportBean();

        for ( int i = 0; i < 10; ++i ) {
            FooSampleReportRow row = new FooSampleReportRow( i );
            report.rowSampleSet1.add( row );
        }
    }

    private Workbook generateWorkbook(String templateFilename, Object report) {
        XLSTransformer transformer =  new XLSTransformer();
        Map map = new HashMap();
        map.put( "report", report );

        InputStream in = null;
        try {
            in = getClass().getResourceAsStream( templateFilename );
            Workbook workbook = transformer.transformXLS( in, map );

            return workbook;
        } catch ( Exception e ) {
            throw new RuntimeException( e );
        } finally {
            try {
                if( in != null ) {
                    in.close();
                }
            } catch ( IOException e ) {
                //suppress
            }
        }
    }

    @Test
    public void jxls_1_0_grouping_sum() throws Exception {
        Workbook actualWorkbook = generateWorkbook(templateFileName, report);
        Sheet resultSheet = actualWorkbook.getSheetAt(0);
        CellsChecker checker = new CellsChecker();
        checker.checkFormulaCell(resultSheet, 12, 0, "SUM(A2:A11)");
        checker.checkFormulaCell(resultSheet, 13, 0, "AVERAGE(A2:A11)");
        checker.checkFormulaCell(resultSheet, 14, 0, "SUM(A2:A11)");
        checker.checkFormulaCell(resultSheet, 15, 0, "AVERAGE(A2:A11)");
        saveWorkbook(actualWorkbook, outputFileName);
    }

    private void saveWorkbook(Workbook resultWorkbook, String fileName) throws IOException {
        OutputStream os = new BufferedOutputStream(new FileOutputStream(fileName));
        resultWorkbook.write(os);
        os.flush();
        os.close();
    }

}
