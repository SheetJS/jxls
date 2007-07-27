package net.sf.jxls;

import junit.framework.TestCase;

import java.util.Set;
import java.util.Iterator;

import net.sf.jxls.formula.Formula;
import net.sf.jxls.formula.CellRef;

/**
 * Test case for Formula class
 */
public class FormulaTest extends TestCase {

    public void testFindRefCells(){
        String formulaValue = "SUM(a1:a10) - D12 + C5 * D10 - 4 + MULT ( B2 : B90 )";
        Formula formula = new Formula( formulaValue );
        Set refCells = formula.findRefCells();
        assertEquals( "Incorrect number of ref cells found", refCells.size(), 7  );
        assertTrue( contains(refCells, "a1") );
        assertTrue( contains(refCells, "a10" ) );
        assertTrue( contains(refCells, "D12" ) );
        assertTrue( contains(refCells, "C5" ) );
        assertTrue( contains(refCells, "D10" ) );
        assertTrue( contains(refCells, "B2" ) );
        assertTrue( contains(refCells, "B90" ) );
    }

    boolean contains(Set refCells, String cellRef){
        for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
            CellRef ref = (CellRef) iterator.next();
            if( ref.toString().equals( cellRef ) ){
                return true;
            }
        }
        return false;
    }

}
