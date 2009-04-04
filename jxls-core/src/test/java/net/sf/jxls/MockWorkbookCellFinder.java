package net.sf.jxls;

import java.util.List;
import java.util.Map;

import net.sf.jxls.controller.WorkbookCellFinderImpl;

import org.apache.poi.hssf.util.CellReference;

/**
 * @author Leonid Vysochyn
 */
public class MockWorkbookCellFinder extends WorkbookCellFinderImpl {
    Map sheetCellsMapping;

    public MockWorkbookCellFinder(Map sheetCellsMapping) {
        this.sheetCellsMapping = sheetCellsMapping;
    }


    public List findCell(String sheetName, String cellName) {
        if (!sheetCellsMapping.containsKey(sheetName)) {
            throw new IllegalArgumentException("Can't find sheet with name " + sheetName);
        }
        Map cellsMapping = (Map) sheetCellsMapping.get(sheetName);
        return (List) cellsMapping.get(cellName);
    }

    public List findCell(String sheetName, int rowNum, int colNum) {
        CellReference cellReference = new CellReference(rowNum, colNum, false, false);
        return findCell(sheetName, cellReference.formatAsString());
    }
}
