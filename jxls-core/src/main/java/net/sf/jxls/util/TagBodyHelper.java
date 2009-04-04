package net.sf.jxls.util;

import java.util.Map;

import net.sf.jxls.tag.Block;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @author Leonid Vysochyn
 */
public class TagBodyHelper {
    protected final Log log = LogFactory.getLog(getClass());

    public static int duplicateDown(HSSFSheet sheet, Block block, int n) {
        if (n > 0) {
            int startRow = block.getEndRowNum() + 1;
            int endRow = sheet.getLastRowNum();
            int numberOfRows = block.getNumberOfRows() * n;
            Util.shiftRows(sheet, startRow, endRow, numberOfRows);
            for (int i = 0; i < n; i++) {
                for (int j = 0; j < block.getNumberOfRows(); j++) {
                    HSSFRow row = sheet.getRow(block.getStartRowNum() + j);
                    HSSFRow newRow = sheet.getRow(block.getEndRowNum() + block.getNumberOfRows() * i + 1 + j);
                    if (row != null) {
                        if (newRow == null) {
                            newRow = sheet.createRow(block.getEndRowNum() + block.getNumberOfRows() * i + 1 + j);
                        }
                        Util.copyRow(sheet, row, newRow);
                    }
                }
            }
            return block.getNumberOfRows() * n;
        }
        return 0;
    }

    public static int duplicateDown(HSSFSheet sheet, Block block, int n, Map formulaCellsToUpdate) {
        if (n > 0) {
            Util.shiftRows(sheet, block.getEndRowNum() + 1, sheet.getLastRowNum(), block.getNumberOfRows() * n);
            for (int i = 0; i < n; i++) {
                for (int j = 0; j < block.getNumberOfRows(); j++) {
                    HSSFRow row = sheet.getRow(block.getStartRowNum() + j);
                    HSSFRow newRow = sheet.getRow(block.getEndRowNum() + block.getNumberOfRows() * i + 1 + j);
                    if (row != null) {
                        if (newRow == null) {
                            newRow = sheet.createRow(block.getEndRowNum() + block.getNumberOfRows() * i + 1 + j);
                        }
                        Util.copyRow(sheet, row, newRow);
                    }
                }
            }
            return block.getNumberOfRows() * n;
        }
        return 0;
    }

    public static int duplicateRight(HSSFSheet sheet, Block block, int n) {
        if (n > 0) {
            Util.shiftCellsRight(sheet, block.getStartRowNum(), block.getEndRowNum(), (short) (block.getEndCellNum() + 1), (short) (block.getNumberOfColumns() * n));
            for (int i = block.getStartRowNum(); i <= block.getEndRowNum(); i++) {
                HSSFRow row = sheet.getRow(i);
                if (row != null) {
                    for (int k = 0; k < n; k++) {
                        for (short j = block.getStartCellNum(); j <= block.getEndCellNum(); j++) {
                            short destCellNum = (short) (block.getEndCellNum() + k * block.getNumberOfColumns() + j - block.getStartCellNum() + 1);
                            HSSFCell destCell = row.getCell(destCellNum);
                            HSSFCell cell = row.getCell(j);
                            if (destCell == null) {
                                destCell = row.createCell(destCellNum);
                            }
                            Util.copyCell(cell, destCell, true);
                            sheet.setColumnWidth(destCellNum, Util.getWidth(sheet, j));
                        }
                    }
                }
            }
            return block.getNumberOfColumns() * n;
        }
        return 0;
    }

    public static void replaceProperty(HSSFSheet sheet, Block block, String oldProperty, String newProperty) {
        for (int i = block.getStartRowNum(); i <= block.getEndRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            replacePropertyInRow(row, oldProperty, newProperty);
        }
    }

    private static void replacePropertyInRow(HSSFRow row, String oldProperty, String newProperty) {
        for (short j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
            HSSFCell cell = row.getCell(j);
            replacePropertyInCell(cell, oldProperty, newProperty);
        }
    }

    private static void replacePropertyInCell(HSSFCell cell, String oldProperty, String newProperty) {
        if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            String cellValue = cell.getRichStringCellValue().getString();
            String newValue = cellValue.replaceAll(oldProperty, newProperty);
            cell.setCellValue(new HSSFRichTextString(newValue));
        }
    }

    public static void removeBorders(HSSFSheet sheet, Block block) {
        HSSFRow rowToDelete = sheet.getRow(block.getStartRowNum());
        deleteRow(sheet, rowToDelete);
        block.setStartRowNum(block.getStartRowNum() + 1);
        deleteRow(sheet, sheet.getRow(block.getEndRowNum()));
        block.setEndRowNum(block.getEndRowNum() - 1);
        shift(sheet, block, -1);
        if (block.getEndRowNum() + 2 < sheet.getLastRowNum()) {
            Util.shiftRows(sheet, block.getEndRowNum() + 3, sheet.getLastRowNum(), -2);
        }
    }

    private static void deleteRow(HSSFSheet sheet, HSSFRow rowToDelete) {
        if (rowToDelete != null) {
            sheet.removeRow(rowToDelete);
        }
    }

    public static void removeLeftRightBorders(HSSFSheet sheet, Block block) {
        HSSFRow row = sheet.getRow(block.getStartRowNum());
        if (row != null) {
            Util.shiftCellsLeft(sheet, block.getStartRowNum(), (short) (block.getStartCellNum() + 1),
                    block.getEndRowNum(), row.getLastCellNum(), (short) 1);
            HSSFCell cellToRemove = row.getCell(row.getLastCellNum());
            clearAndRemoveCell(row, cellToRemove);
            Util.shiftCellsLeft(sheet, block.getStartRowNum(), block.getEndCellNum(), block.getEndRowNum(), row.getLastCellNum(), (short) 1);
            HSSFCell cell = cellToRemove;
            clearAndRemoveCell(row, cell);
            block.setEndCellNum((short) (block.getEndCellNum() - 2));
        }
    }

    private static void clearAndRemoveCell(HSSFRow row, HSSFCell cellToRemove) {
        clearCell(cellToRemove);
        if (cellToRemove != null) {
            row.removeCell(cellToRemove);
        }
    }

    public static void shift(HSSFSheet sheet, Block block, int n) {
        Util.shiftRows(sheet, block.getStartRowNum(), block.getEndRowNum(), n);
        block.setStartRowNum(block.getStartRowNum() + n);
        block.setEndRowNum(block.getEndRowNum() + n);

    }


    public static void removeRowCells(HSSFSheet sheet, HSSFRow row, short startCellNum, short endCellNum) {
        clearRowCells(row, startCellNum, endCellNum);
        Util.shiftCellsLeft(sheet, row.getRowNum(), (short) (endCellNum + 1), row.getRowNum(), row.getLastCellNum(), (short) (endCellNum - startCellNum + 1));
        clearRowCells(row, (short) (row.getLastCellNum() - (endCellNum - startCellNum)), row.getLastCellNum());
    }

    public static void removeBodyRows(HSSFSheet sheet, Block block) {
        for (int i = 0; i < block.getNumberOfRows(); i++) {
            HSSFRow row = sheet.getRow(block.getStartRowNum() + i);
            removeMergedRegions(sheet, row);
            deleteRow(sheet, row);
        }
        Util.shiftRows(sheet, block.getEndRowNum() + 1, sheet.getLastRowNum(), -block.getNumberOfRows());
    }

    private static void removeMergedRegions(HSSFSheet sheet, HSSFRow row) {
        if (row != null) {
            int i = row.getRowNum();
            for (short j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                Util.removeMergedRegion(sheet, i, j);
            }
        }
    }

    static void clearRow(HSSFRow row) {
        if (row != null) {
            for (short i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
                HSSFCell cell = row.getCell(i);
                clearCell(cell);
            }
        }
    }

    static void clearRowCells(HSSFRow row, short startCell, short endCell) {
        if (row != null) {
            for (short i = startCell; i <= endCell; i++) {
                HSSFCell cell = row.getCell(i);
                if (cell != null) {
                    row.removeCell(cell);
                }
                row.createCell(i);
            }
        }
    }

    static void clearCell(HSSFCell cell) {
        if (cell != null) {
            cell.setCellValue(new HSSFRichTextString(""));
            cell.setCellType(HSSFCell.CELL_TYPE_BLANK);
        }
    }

    public static void adjustFormulas(HSSFWorkbook hssfWorkbook, HSSFSheet hssfSheet, Block body) {
        for (int i = body.getStartRowNum(); i <= body.getEndRowNum(); i++) {
            HSSFRow row = hssfSheet.getRow(i);
            adjustFormulas(row);
        }
    }

    private static void adjustFormulas(HSSFRow row) {
        if (row != null) {
            for (short i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
                HSSFCell cell = row.getCell(i);
                if (cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING && cell.getRichStringCellValue().getString().matches("\\$\\[.*?\\]")) {
                    String cellValue = cell.getRichStringCellValue().getString();
                    String[] parts = cellValue.split("\\$\\[.*?\\]");
                    String newCellValue = parts[0];
                    newCellValue = newCellValue.replaceAll("#", Integer.toString(row.getRowNum() + 1));

                    cell.setCellValue(new HSSFRichTextString(newCellValue));
                }
            }
        }
    }
}
