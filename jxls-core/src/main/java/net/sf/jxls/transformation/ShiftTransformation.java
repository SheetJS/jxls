package net.sf.jxls.transformation;

import net.sf.jxls.formula.CellRef;
import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;
import org.apache.poi.hssf.util.CellReference;

import java.util.ArrayList;
import java.util.List;

/**
 * Defines simple shift transformation
 * @author Leonid Vysochyn
 */
public class ShiftTransformation extends BlockTransformation {
    int rowShift, colShift;
    int rowNum;
    int colNum;
    private CellReference cellReference;
    private List cells = new ArrayList();

    public ShiftTransformation(Block block, int rowShift, int colShift) {
        super(block);
        this.rowShift = rowShift;
        this.colShift = colShift;
    }

    public Block getBlockAfterTransformation() {
        return null;
    }

    public List transformCell(Point p) {
        cells.clear();
        if( block.contains( p ) || (block.isAbove( p ) && rowShift != 0) || (block.isToLeft( p ) && colShift != 0)){
            Point newPoint = p.shift( rowShift, colShift );
            cells.add( newPoint );
        }else{
            cells.add( p );
        }
        return cells;
    }

    public List transformCell(String sheetName, CellRef cellRef) {
        cells.clear();
        String refSheetName = cellRef.getSheetName();
        if( block.getSheet().getSheetName().equalsIgnoreCase( refSheetName ) || (cellRef.getSheetName() == null && block.getSheet().getSheetName().equalsIgnoreCase( sheetName ))){
            if( block.contains( cellRef.getRowNum(), cellRef.getColNum() ) || (block.getEndRowNum() < cellRef.getRowNum() && rowShift != 0)
                    || (block.getEndCellNum() < cellRef.getColNum() && colShift != 0)){
                rowNum = cellRef.getRowNum() + rowShift;
                colNum = cellRef.getColNum() + colShift;
                // todo: remove this check
                if( colNum < 0 ){
                    colNum = 0;
                }
                cellReference = new CellReference( rowNum, colNum );
                if( cellRef.getSheetName() != null ){
                    cells.add( cellRef.getSheetName() + "!" + cellReference.toString());
                }else{
                    cells.add( cellReference.toString() );
                }
            }
        }
        return cells;
    }

    public boolean equals(Object obj) {
        if( obj != null && obj instanceof ShiftTransformation ){
            ShiftTransformation st = (ShiftTransformation) obj;
            return ( super.equals( obj ) && rowShift == st.rowShift && colShift == st.colShift);
        }else{
            return false;
        }
    }

    public int hashCode() {
        int result = super.hashCode();
        result = 29 * result + rowShift;
        result = 29 * result + colShift;
        return result;
    }

    public String toString() {
        return "ShiftTransformation: {" + super.toString() + ", shift=(" + rowShift + ", " + colShift + ")}";
    }
}
