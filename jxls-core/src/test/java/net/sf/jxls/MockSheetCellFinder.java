package net.sf.jxls;

import org.apache.poi.hssf.util.CellReference;
import net.sf.jxls.controller.SheetCellFinder;

import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class MockSheetCellFinder implements SheetCellFinder {

    Map cellsMapping;

    public MockSheetCellFinder(Map cellsMapping) {
        this.cellsMapping = cellsMapping;
    }

    public List findCell(String cellName) {
        return (List) cellsMapping.get( cellName );
    }

    public List findCell(int rowNum, int colNum) {
        CellReference cellReference = new CellReference( rowNum, colNum );
        return (List) cellsMapping.get( cellReference.toString() );
    }

    public List findCell(String sheetName, String cellName) {
        return (List) cellsMapping.get( cellName );
    }

    public List findCell(String sheetName, int rowNum, int colNum) {
        CellReference cellReference = new CellReference( rowNum, colNum );
        return (List) cellsMapping.get( cellReference.toString() );
    }
}
