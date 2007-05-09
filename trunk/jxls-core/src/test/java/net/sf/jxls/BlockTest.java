package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.tag.Block;

/**
 * @author Leonid Vysochyn
 */
public class BlockTest extends TestCase {
    public void testEquals(){
        Block b1 = new Block(1, (short) 2, 3, (short)4);
        Block b2 = new Block(1, (short) 2, 3, (short)4);
        assertTrue( b1.equals( b2 ));
        b2 = new Block(0, (short) 2, 3, (short)4);
        assertFalse( b1.equals( b2 ));
        b2 = new Block(1, (short) 3, 3, (short)4);
        assertFalse( b1.equals( b2 ));
        b2 = new Block(1, (short) 2, 5, (short)4);
        assertFalse( b1.equals( b2 ));
        b2 = new Block(1, (short) 2, 3, (short)5);
        assertFalse( b1.equals( b2 ));
    }
}
