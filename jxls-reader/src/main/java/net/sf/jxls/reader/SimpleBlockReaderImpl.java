package net.sf.jxls.reader;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class SimpleBlockReaderImpl extends BaseBlockReader implements SimpleBlockReader {
    protected final Log log = LogFactory.getLog(getClass());

    List beanCellMappings = new ArrayList();
    SectionCheck sectionCheck;

    static {
        ReaderConfig.getInstance();
    }

    public SimpleBlockReaderImpl() {
    }

    public SimpleBlockReaderImpl(int startRow, int endRow, List beanCellMappings) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.beanCellMappings = beanCellMappings;
    }

    public SimpleBlockReaderImpl(int startRow, int endRow) {
        this.startRow = startRow;
        this.endRow = endRow;
    }

    public XLSReadStatus read(XLSRowCursor cursor, Map beans) {
        readStatus.clear();
        final int currentRowNum = cursor.getCurrentRowNum();
        final int rowShift = currentRowNum - startRow;
        BeanCellMapping mapping;
        for (Iterator iterator = beanCellMappings.iterator(); iterator.hasNext();) {
            mapping = (BeanCellMapping) iterator.next();
            try {
                String dataString = readCellString(cursor.getSheet(), mapping.getRow() + rowShift, mapping.getCol());
                mapping.populateBean(dataString, beans);
            } catch (Exception e) {
                String message = "Can't read cell " + getCellName(mapping, rowShift) + " on " + cursor.getSheetName() + " spreadsheet";
                readStatus.addMessage(new XLSReadMessage(message, e));
                if (ReaderConfig.getInstance().isSkipErrors()) {
                    if (log.isWarnEnabled()) {
                        log.warn(message);
                    }
                } else {
                    readStatus.setStatusOK(false);
                    throw new XLSDataReadException(getCellName(mapping, rowShift), "Can't read cell " + getCellName(mapping, rowShift) + " on " + cursor.getSheetName() + " spreadsheet", readStatus);
                }
            }
        }
        cursor.setCurrentRowNum(endRow + rowShift);
        return readStatus;
    }

    private String readCellString(Sheet sheet, int rowNum, short cellNum) {
        Cell cell = getCell(sheet, rowNum, cellNum);
        return getCellString(cell);
    }

    private String getCellString(Cell cell) {
        String dataString = null;
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    dataString = cell.getRichStringCellValue().getString();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    dataString = readNumericCell(cell);
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    dataString = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                case Cell.CELL_TYPE_ERROR:
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    // attempt to read formula cell as numeric cell
                    try{
                    dataString = readNumericCell(cell);
                    }catch(Exception e1){
                        log.info("Failed to read formula cell as numeric. Next to try as string. Cell=" + cell.toString());
                        try{
                            dataString = cell.getRichStringCellValue().getString();
                            log.info("Successfully read formula cell as string. Value=" + dataString);
                        }catch(Exception e2){
                            log.warn("Failed to read formula cell as numeric or string. Cell=" + cell.toString());
                        }
                    }

                    break;
                default:
                    break;
            }
        }
        return dataString;
    }

    private String readNumericCell(Cell cell) {
        double value;
        String dataString = null;
        value = cell.getNumericCellValue();
        if (((int) value) == value) {
            dataString = Integer.toString((int) value);
        } else {
            dataString = Double.toString(cell.getNumericCellValue());
        }
        return dataString;
    }

    private String getCellName(BeanCellMapping mapping, int rowShift) {
        CellReference currentCellRef = new CellReference(mapping.getRow() + rowShift, mapping.getCol(), false, false);
        return currentCellRef.formatAsString();
    }


    public SectionCheck getLoopBreakCondition() {
        return sectionCheck;
    }

    public void setLoopBreakCondition(SectionCheck sectionCheck) {
        this.sectionCheck = sectionCheck;
    }

    public void addMapping(BeanCellMapping mapping) {
        beanCellMappings.add(mapping);
    }

    public List getMappings() {
        return beanCellMappings;
    }

    private Cell getCell(Sheet sheet, int rowNum, int cellNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return null;
        }
        return row.getCell(cellNum);
    }

}
