package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.tag.Block;
import net.sf.jxls.transformation.DuplicateTransformation;
import net.sf.jxls.transformer.Sheet;

/**
 * @author Leonid Vysochyn
 */
public class DuplicateTransformationTest extends TestCase {
    public void testEqualsTrue(){
        Sheet sheet = new Sheet();
        sheet.setSheetName( "Sheet 1");
        Block b1 = new Block(1, (short)2, 3, (short) 4);
        Block b2 = new Block(1, (short)2, 3, (short) 4);
        b1.setSheet( sheet );
        b2.setSheet( sheet );

        DuplicateTransformation dt1 = new DuplicateTransformation(b1, 1);
        DuplicateTransformation dt2 = new DuplicateTransformation(b2, 1);
        assertTrue (dt1.equals(dt2));
        dt1 = new DuplicateTransformation(null, 2);
        dt2 = new DuplicateTransformation(null, 2);
        assertTrue( dt1.equals(dt2) );
    }
    public void testEqualsFalse(){
        Sheet sheet = new Sheet();
        sheet.setSheetName( "Sheet 1");
        Block b1 = new Block(1, (short)2, 3, (short) 4);
        Block b2 = new Block(1, (short)2, 3, (short) 5);
        b1.setSheet( sheet );
        b2.setSheet( sheet );
        DuplicateTransformation dt1 = new DuplicateTransformation(b1, 1);
        DuplicateTransformation dt2 = new DuplicateTransformation(b2, 1);
        assertFalse (dt1.equals(dt2));
        dt2 = new DuplicateTransformation(b1, 0);
        assertFalse (dt1.equals(dt2));
        dt1 = new DuplicateTransformation(null, 1);
        assertFalse( dt1.equals(dt2) );
    }
}
