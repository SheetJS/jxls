package net.sf.jxls.transformation;

import net.sf.jxls.transformation.BlockTransformation;
import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;
import net.sf.jxls.formula.CellRef;

import java.util.List;
import java.util.ArrayList;

/**
 * Remove transformation
 * @author Leonid Vysochyn
 */
public class RemoveTransformation extends BlockTransformation {

    public RemoveTransformation(Block block) {
        super(block);
    }

    public Block getBlockAfterTransformation() {
        return null;  //To change body of implemented methods use File | Settings | File Templates.
    }


    public List transformCell(Point p) {
        List cells = new ArrayList();
        if( block.contains( p ) ){
            cells.add( null );
        }else{
            cells.add( p );
        }
        return cells;
    }

    public List transformCell(String sheetName, CellRef cellRef) {
        return null;
    }

    public String toString() {
        return "RemoveTransformation: {" + super.toString() + "}";
    }
}
