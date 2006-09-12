package net.sf.jxls.controller;

import org.apache.poi.hssf.util.CellReference;
import net.sf.jxls.controller.SheetCellFinder;

import java.util.List;
import java.util.ArrayList;

import net.sf.jxls.controller.MatrixCell;

/**
 * Implementation of SheetCellFinder interface. This implementation is based on a virtual object matrix.
 * @author Leonid Vysochyn
 */
public class SheetCellFinderImpl implements SheetCellFinder {

    TransformationMatrix transformationMatrix;

    public SheetCellFinderImpl(TransformationMatrix transformationMatrix) {
        this.transformationMatrix = transformationMatrix;
    }

    public List findCell(String cellName){
        CellReference cellReference = new CellReference( cellName );
        int colNum = cellReference.getCol();
        int rowNum = cellReference.getRow();
        List matrixCells = findCell( rowNum, colNum );
        return convertToCellNames( matrixCells );
    }

    public List findCell(int rowNum, int colNum){
        List matrixCells = transformationMatrix.findMappedCells( rowNum, colNum );
        return convertToCellNames( matrixCells );
    }

    private List convertToCellNames(List matrixCells) {
        List stringCells = new ArrayList();
        for (int i = 0; i < matrixCells.size(); i++) {
            MatrixCell cell = (MatrixCell) matrixCells.get(i);
            stringCells.add( new CellReference(cell.getRowNum(), cell.getColNum()).toString() );
        }
        return stringCells;
    }


}
