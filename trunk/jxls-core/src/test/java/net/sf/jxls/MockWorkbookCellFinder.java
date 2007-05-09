package net.sf.jxls;

import org.apache.poi.hssf.util.CellReference;
import net.sf.jxls.controller.WorkbookCellFinderImpl;

import java.util.Map;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class MockWorkbookCellFinder extends WorkbookCellFinderImpl {
    Map sheetCellsMapping;

    public MockWorkbookCellFinder(Map sheetCellsMapping) {
        this.sheetCellsMapping = sheetCellsMapping;
    }



    public List findCell(String sheetName, String cellName) {
        if( !sheetCellsMapping.containsKey( sheetName ) ){
            throw new IllegalArgumentException("Can't find sheet with name " + sheetName);
        }
        Map cellsMapping = (Map) sheetCellsMapping.get( sheetName );
        return (List) cellsMapping.get( cellName );
    }

    public List findCell(String sheetName, int rowNum, int colNum) {
        CellReference cellReference = new CellReference( rowNum, colNum );
        return findCell( sheetName, cellReference.toString() );
    }
}
