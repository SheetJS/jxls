package net.sf.jxls;

import junit.framework.Assert;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class CellsChecker extends Assert {

    Map propertyMap = new HashMap();

    public CellsChecker() {
    }

    boolean ignoreStyle = false;
    boolean ignoreFirstLastCellNums = false;

    public CellsChecker(Map propertyMap) {
        this.propertyMap = propertyMap;
    }

    public CellsChecker(Map propertyMap, boolean ignoreStyle) {
        this.propertyMap = propertyMap;
        this.ignoreStyle = ignoreStyle;
    }


    public boolean isIgnoreFirstLastCellNums() {
        return ignoreFirstLastCellNums;
    }

    public void setIgnoreFirstLastCellNums(boolean ignoreFirstLastCellNums) {
        this.ignoreFirstLastCellNums = ignoreFirstLastCellNums;
    }

    void checkSection(Sheet srcSheet, Sheet destSheet, int srcRowNum, int destRowNum, short fromCellNum, short toCellNum, int numberOfRows, boolean ignoreHeight, boolean ignoreNullRows) {
        for (int i = 0; i < numberOfRows; i++) {
            Row sourceRow = srcSheet.getRow(srcRowNum + i);
            Row destRow = destSheet.getRow(destRowNum + i);
            if (!ignoreNullRows) {
                assertTrue("Null Row problem found", (sourceRow != null && destRow != null) || (sourceRow == null && destRow == null));
                if (sourceRow != null) {
                    if (!ignoreHeight) {
                        assertEquals("Row height is not the same", sourceRow.getHeight(), destRow.getHeight());
                    }
                    checkCells(sourceRow, destRow, fromCellNum, toCellNum);
                }
            } else {
                if (!ignoreHeight) {
                    assertEquals("Row height is not the same", sourceRow.getHeight(), destRow.getHeight());
                }
                if (sourceRow == null && destRow != null) {
                    checkEmptyCells(destRow, fromCellNum, toCellNum);
                }
                if (destRow == null && sourceRow != null) {
                    checkEmptyCells(sourceRow, fromCellNum, toCellNum);
                }
                if (sourceRow != null && destRow != null) {
                    checkCells(sourceRow, destRow, fromCellNum, toCellNum);
                }
            }
        }
    }

    void checkRow(Sheet sheet, int rowNum, int startCellNum, int endCellNum, Object[] values){
        Row row  = sheet.getRow(rowNum);
        if( row != null){
            for(int i = startCellNum; i<=endCellNum; i++){
                Cell cell = row.getCell(i);
                if( cell != null ){
                    Object cellValue = getCellValue(cell, values[i]);
                    assertEquals("Result cell values incorrect in row=" + row + ", cell=" + i, values[i], cellValue);
                }else{
                    fail("Cell is null");
                }
            }
        }else{
            fail("Row is null");
        }
    }

    void checkCell(Sheet sheet, int rowNum, int cellNum, Object value){
        Row row  = sheet.getRow(rowNum);
        if( row != null){
            Cell cell = row.getCell(cellNum);
            if( cell != null ){
                Object cellValue = getCellValue(cell, value);
                assertEquals("Result cell values incorrect in row=" + row + ", cell=" + cellNum, value, cellValue);
            }else{
                fail("Cell is null");
            }
        }else{
            fail("Row is null");
        }
    }

    private void checkEmptyCells(Row destRow, short fromCellNum, short toCellNum) {
        if (destRow != null) {
            for (short i = fromCellNum; i <= toCellNum; i++) {
                assertNull("Cell " + i + " in " + destRow.getRowNum() + " row is not null", destRow.getCell(i));
            }
        }
    }

    void checkListCells(Sheet srcSheet, int srcRowNum, Sheet sheet, int startRowNum, short cellNum, Object[] values) {
        Row srcRow = srcSheet.getRow(srcRowNum);
        Cell srcCell = srcRow.getCell(cellNum);
        for (int i = 0; i < values.length; i++) {
            Row row = sheet.getRow(startRowNum + i);
            Cell cell = row.getCell(cellNum);
            Object cellValue = getCellValue(cell, values[i]);
            assertEquals("List property cell is incorrect", values[i], cellValue);
            checkCellStyle(srcCell.getCellStyle(), cell.getCellStyle());
        }
    }

    void checkFixedListCells(Sheet srcSheet, int srcRowNum, Sheet destSheet, int startRowNum, short cellNum, Object[] values) {
        for (int i = 0; i < values.length; i++) {
            Row srcRow = srcSheet.getRow(srcRowNum);
            Cell srcCell = srcRow.getCell(cellNum);
            Row destRow = destSheet.getRow(startRowNum + i);
            Cell destCell = destRow.getCell(cellNum);
            Object cellValue = getCellValue(destCell, values[i]);
            assertEquals("List property cell is incorrect", values[i], cellValue);
            checkCellStyle(srcCell.getCellStyle(), destCell.getCellStyle());
        }
    }

    void checkFormulaCell(Sheet sheet, int rowNum, int cellNum, String formula){
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(cellNum);
        assertEquals("Result Cell is not a formula", cell.getCellType(), Cell.CELL_TYPE_FORMULA);
        assertEquals("Formula is incorrect", formula, cell.getCellFormula());
    }

    void checkFormulaCell(Sheet srcSheet, int srcRowNum, Sheet destSheet, int destRowNum, short cellNum, String formula) {
        Row srcRow = srcSheet.getRow(srcRowNum);
        Cell srcCell = srcRow.getCell(cellNum);
        Row destRow = destSheet.getRow(destRowNum);
        Cell destCell = destRow.getCell(cellNum);
        checkCellStyle(srcCell.getCellStyle(), destCell.getCellStyle());
        assertEquals("Result Cell is not a formula", destCell.getCellType(), Cell.CELL_TYPE_FORMULA);
        assertEquals("Formula is incorrect", formula, destCell.getCellFormula());
    }

    void checkFormulaCell(Sheet srcSheet, int srcRowNum, Sheet destSheet, int destRowNum, short cellNum, String formula, boolean ignoreCellStyle) {
        Row srcRow = srcSheet.getRow(srcRowNum);
        Cell srcCell = srcRow.getCell(cellNum);
        Row destRow = destSheet.getRow(destRowNum);
        Cell destCell = destRow.getCell(cellNum);
        if (!ignoreCellStyle) {
            checkCellStyle(srcCell.getCellStyle(), destCell.getCellStyle());
        }
        assertEquals("Result Cell is not a formula", destCell.getCellType(), Cell.CELL_TYPE_FORMULA);
        assertEquals("Formula is incorrect", formula, destCell.getCellFormula());
    }

    void checkRows(Sheet sourceSheet, Sheet destSheet, int sourceRowNum, int destRowNum, int numberOfRows, boolean checkRowHeight) {
        for (int i = 0; i < numberOfRows; i++) {
            Row sourceRow = sourceSheet.getRow(sourceRowNum + i);
            Row destRow = destSheet.getRow(destRowNum + i);
            assertTrue("Null Row problem found", (sourceRow != null && destRow != null) || (sourceRow == null && destRow == null));
            if (sourceRow != null && destRow != null) {
                if (!ignoreFirstLastCellNums) {
                    assertEquals("First Cell Numbers differ in source and result row", sourceRow.getFirstCellNum(), destRow.getFirstCellNum());
                }
                assertEquals("Physical Number Of Cells differ in source and result row", sourceRow.getPhysicalNumberOfCells(), destRow.getPhysicalNumberOfCells());
                if( checkRowHeight ){
                    assertEquals("Row height is not the same for srcRow = " + sourceRow.getRowNum() + ", destRow = " + destRow.getRowNum(),
                            sourceRow.getHeight(), destRow.getHeight());
                }
                checkCells(sourceRow, destRow, sourceRow.getFirstCellNum(), sourceRow.getLastCellNum());
            }
        }
    }

    private void checkCells(Row sourceRow, Row resultRow, short startCell, short endCell) {
		 if (startCell >= 0 && endCell >= 0) {
			  for (short i = startCell; i <= endCell; i++) {
					Cell sourceCell = sourceRow.getCell(i);
					Cell resultCell = resultRow.getCell(i);
					assertTrue("Null cell problem found", (sourceCell != null && resultCell != null) || (sourceCell == null && resultCell == null));
					if (sourceCell != null) {
						 checkCells(sourceCell, resultCell);
					}
			  }
		 }
    }

    void checkCells(Sheet srcSheet, Sheet destSheet, int srcRowNum, short srcCellNum, int destRowNum, short destCellNum, boolean checkCellWidth) {
        Row srcRow = srcSheet.getRow(srcRowNum);
        Row destRow = destSheet.getRow(destRowNum);
        assertEquals("Row height is not the same", srcRow.getHeight(), destRow.getHeight());
        Cell srcCell = srcRow.getCell(srcCellNum);
        Cell destCell = destRow.getCell(destCellNum);
        assertTrue("Null cell problem found", (srcCell != null && destCell != null) || (srcCell == null && destCell == null));
        if (srcCell != null && destCell != null) {
            checkCells(srcCell, destCell);
        }
        if (checkCellWidth) {
            assertEquals("Cell Widths are different", getWidth(srcSheet, srcCellNum), getWidth(destSheet, destCellNum));
        }
    }

    static int getWidth(Sheet sheet, int col) {
        int width = sheet.getColumnWidth(col);
        if (width == sheet.getDefaultColumnWidth()) {
            width = (short) (width * 256);
        }
        return width;
    }


    private void checkCells(Cell sourceCell, Cell destCell) {
        checkCellValue(sourceCell, destCell);
        checkCellStyle(sourceCell.getCellStyle(), destCell.getCellStyle());
    }

    private void checkCellStyle(CellStyle sourceStyle, CellStyle destStyle) {
        if (!ignoreStyle) {
            assertEquals(sourceStyle.getAlignment(), destStyle.getAlignment());
            assertEquals(sourceStyle.getBorderBottom(), destStyle.getBorderBottom());
            assertEquals(sourceStyle.getBorderLeft(), destStyle.getBorderLeft());
            assertEquals(sourceStyle.getBorderRight(), destStyle.getBorderRight());
            assertEquals(sourceStyle.getBorderTop(), destStyle.getBorderTop());
            assertEquals(sourceStyle.getBottomBorderColor(), sourceStyle.getBottomBorderColor());
            assertEquals(sourceStyle.getFillBackgroundColor(), destStyle.getFillBackgroundColor());
            assertEquals(sourceStyle.getFillForegroundColor(), sourceStyle.getFillForegroundColor());
            assertEquals(sourceStyle.getFillPattern(), destStyle.getFillPattern());
            assertEquals(sourceStyle.getHidden(), destStyle.getHidden());
            assertEquals(sourceStyle.getIndention(), destStyle.getIndention());
            assertEquals(sourceStyle.getLeftBorderColor(), destStyle.getLeftBorderColor());
            assertEquals(sourceStyle.getLocked(), destStyle.getLocked());
            assertEquals(sourceStyle.getRightBorderColor(), destStyle.getRightBorderColor());
            assertEquals(sourceStyle.getRotation(), destStyle.getRotation());
            assertEquals(sourceStyle.getTopBorderColor(), destStyle.getTopBorderColor());
            assertEquals(sourceStyle.getVerticalAlignment(), destStyle.getVerticalAlignment());
            assertEquals(sourceStyle.getWrapText(), destStyle.getWrapText());
        }
    }

    private void checkCellValue(Cell sourceCell, Cell destCell) {
        switch (sourceCell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                if (propertyMap.containsKey(sourceCell.getRichStringCellValue().getString())) {
                    assertEquals("Property value was set incorrectly", propertyMap.get(sourceCell.getRichStringCellValue().getString()), getCellValue(destCell, propertyMap.get(sourceCell.getRichStringCellValue().getString())));
                } else {
                    assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                    assertEquals("Cell values are not the same", sourceCell.getRichStringCellValue().getString(), destCell.getRichStringCellValue().getString());
                }
                break;
            case Cell.CELL_TYPE_NUMERIC:
                assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                assertTrue("Cell values are not the same", sourceCell.getNumericCellValue() == destCell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                assertEquals("Cell values are not the same", sourceCell.getBooleanCellValue(), destCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                assertEquals("Cell values are not the same", sourceCell.getErrorCellValue(), destCell.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                assertEquals("Cell values are not the same", sourceCell.getCellFormula(), destCell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK:
                assertEquals("Cell type is not the same", sourceCell.getCellType(), destCell.getCellType());
                break;
            default:
                fail("Unknown cell type, code=" + sourceCell.getCellType() + ", value=" + sourceCell.getRichStringCellValue().getString());
                break;
        }
    }

    private Object getCellValue(Cell cell, Object obj) {
        Object value = null;
        if (obj instanceof String) {
            value = cell.getRichStringCellValue().getString();
        } else if (obj instanceof Double) {
            value = new Double(cell.getNumericCellValue());
        } else if (obj instanceof BigDecimal) {
            value = new BigDecimal(cell.getNumericCellValue());
        } else if (obj instanceof Integer) {
            value = new Integer((int) cell.getNumericCellValue());
        } else if (obj instanceof Float) {
            value = new Float(cell.getNumericCellValue());
        } else if (obj instanceof Date) {
            value = cell.getDateCellValue();
        } else if (obj instanceof Calendar) {
            Calendar c = Calendar.getInstance();
            c.setTime(cell.getDateCellValue());
            value = c;
        } else if (obj instanceof Boolean) {
            if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                value = (cell.getBooleanCellValue()) ? Boolean.TRUE : Boolean.FALSE;
            } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                value = Boolean.valueOf(cell.getRichStringCellValue().getString());
            } else {
                value = Boolean.FALSE;
            }
        }
        return value;
    }

}
