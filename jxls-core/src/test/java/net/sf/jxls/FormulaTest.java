package net.sf.jxls;

import junit.framework.TestCase;

import java.util.Set;

import net.sf.jxls.formula.Formula;

/**
 * Test case for Formula class
 */
public class FormulaTest extends TestCase {

    public void testFindRefCells(){
        String formulaValue = "SUM(a1:a10) - D12 + C5 * D10 - 4 + MULT ( B2 : B90 )";
        Formula formula = new Formula( formulaValue );
        Set refCells = formula.findRefCells();
        assertEquals( "Incorrect number of ref cells found", refCells.size(), 7  );
        assertTrue( refCells.contains("a1" ) );
        assertTrue( refCells.contains("a10" ) );
        assertTrue( refCells.contains("D12" ) );
        assertTrue( refCells.contains("C5" ) );
        assertTrue( refCells.contains("D10" ) );
        assertTrue( refCells.contains("B2" ) );
        assertTrue( refCells.contains("B90" ) );
    }

    public void testAdjust(){

    }
}
