package net.sf.jxls;

import junit.framework.TestCase;

import java.util.Map;
import java.util.HashMap;
import java.util.Arrays;

import net.sf.jxls.formula.Formula;
import net.sf.jxls.formula.CommonFormulaResolver;
import net.sf.jxls.formula.FormulaResolver;
import net.sf.jxls.controller.WorkbookCellFinder;

/**
 * @author Leonid Vysochyn
 */
public class FormulaResolverTest extends TestCase {

    public void testResolve(){
        String[] a1Mapped = {"d15", "d16", "d18"};
        String[] b10Mapped = {"c30", "C31", "C32", "c33"};
        String[] an20Mapped = {"bk20", "bl20", "BM20", "BN20"};
        String[] az10Mapped = {"CK10", "CL10", "CN10", "CM10"};
        String[] c5Mapped = {"C5"};
        String[] d15Mapped = {"E25"};
        String[] z5Mapped = {"k1", "k2", "k3"};
        Map cellsMapping = new HashMap();
        Map sheetCellsMapping = new HashMap();
        cellsMapping.put( "A1", Arrays.asList( a1Mapped ) );
        cellsMapping.put( "B10", Arrays.asList( b10Mapped ) );
        cellsMapping.put( "AN20", Arrays.asList( an20Mapped ) );
        sheetCellsMapping.put("TestSheet", cellsMapping);
        cellsMapping = new HashMap();
        cellsMapping.put( "C5", Arrays.asList( c5Mapped ) );
        cellsMapping.put( "D15", Arrays.asList( d15Mapped ) );
        cellsMapping.put( "AZ10", Arrays.asList( az10Mapped ) );
        cellsMapping.put( "Z5", Arrays.asList( z5Mapped ) );
        sheetCellsMapping.put("&#_-Bum 25 Sheet_", cellsMapping);
        WorkbookCellFinder cellFinder = new MockWorkbookCellFinder( sheetCellsMapping );
        FormulaResolver formulaResolver = new CommonFormulaResolver();
        String templateFormula = " SUM( B10) - '&#_-Bum 25 Sheet_'!D15 + '&#_-Bum 25 Sheet_'!C5 * (A1) - 4 + MULT " +
                "( AN20 )  * sum('&#_-Bum 25 Sheet_'!AZ10) - SUM('&#_-Bum 25 Sheet_'!Z5)";
        String expectedFormula = " SUM( C30:C33) - '&#_-Bum 25 Sheet_'!E25 + '&#_-Bum 25 Sheet_'!C5 * (D15,D16,D18) - " +
                "4 + MULT ( BK20:BN20 )  * sum('&#_-Bum 25 Sheet_'!CK10,'&#_-Bum 25 Sheet_'!CL10,'&#_-Bum 25 Sheet_'!CN10,'&#_-Bum 25 Sheet_'!CM10)" +
                " - SUM('&#_-Bum 25 Sheet_'!K1:K3)";
        MockSheet mockSheet = new MockSheet();
        mockSheet.setSheetName( "TestSheet" );
        Formula formula = new Formula( templateFormula );
        formula.setSheet( mockSheet );
        String resultFormula = formulaResolver.resolve( formula, cellFinder );
//        assertEquals("Resolved formula is incorrect", expectedFormula, resultFormula );
    }


}
