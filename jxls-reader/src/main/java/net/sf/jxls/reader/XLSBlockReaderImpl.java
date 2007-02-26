package net.sf.jxls.reader;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class XLSBlockReaderImpl implements XLSBlockReader {
    protected final Log log = LogFactory.getLog(getClass());

    int startRow;
    int endRow;

    List beanCellMappings = new ArrayList();
    SectionCheck sectionCheck;


    public XLSBlockReaderImpl() {
    }

    public XLSBlockReaderImpl(int startRow, int endRow, List beanCellMappings) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.beanCellMappings = beanCellMappings;
    }

    public void read(XLSRowCursor cursor, Map beans)  {
        int currentRowNum = cursor.getCurrentRowNum();
        int shift = currentRowNum - startRow;
        BeanCellMapping lastProcessedCellMapping = null;
        try {
            for (Iterator iterator = beanCellMappings.iterator(); iterator.hasNext();) {
                BeanCellMapping mapping = (BeanCellMapping) iterator.next();
                HSSFCell cell = getCell( cursor.getSheet(), mapping.getRow() + shift, mapping.getCol());
                Object value = getCellValue( cell );
                mapping.populateBean( value, beans );
                lastProcessedCellMapping = mapping;
            }
        } catch (Exception e) {
            throw new XLSDataReadException( "Can't parse XLS. Last processed cell was " + (lastProcessedCellMapping==null?null:lastProcessedCellMapping.getCellName()), e );
        }
        cursor.setCurrentRowNum( endRow + shift );
    }


    public SectionCheck getLoopBreakCondition() {
        return sectionCheck;
    }

    public void setLoopBreakCondition(SectionCheck sectionCheck) {
        this.sectionCheck = sectionCheck;
    }

    public void addBlockReader(XLSBlockReader reader) {
    }

    public List getBlockReaders() {
        return null;  
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public void addMapping(BeanCellMapping mapping) {
        beanCellMappings.add( mapping );
    }

    private Object getCellValue(HSSFCell cell) {
        if( cell == null ){
            return null;
        }
        Object value = null;
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                value = new Double(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                value = (cell.getBooleanCellValue()) ? Boolean.TRUE : Boolean.FALSE;
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                break;
            default:
                break;
        }
        return value;
    }

    private HSSFCell getCell(HSSFSheet sheet, int rowNum, short cellNum){
        HSSFRow row = sheet.getRow( rowNum );
        if( row == null ){
            return null;
        }else{
            return row.getCell( cellNum );
        }
    }

}
