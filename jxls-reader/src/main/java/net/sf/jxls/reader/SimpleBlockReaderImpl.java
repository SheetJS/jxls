package net.sf.jxls.reader;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.CellReference;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class SimpleBlockReaderImpl implements SimpleBlockReader {
    protected final Log log = LogFactory.getLog(getClass());

    int startRow;
    int endRow;

    List beanCellMappings = new ArrayList();
    SectionCheck sectionCheck;


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

    public void read(XLSRowCursor cursor, Map beans)  {
        int currentRowNum = cursor.getCurrentRowNum();
        int shift = currentRowNum - startRow;
        CellReference currentCellRef = null;
        try {
            for (Iterator iterator = beanCellMappings.iterator(); iterator.hasNext();) {
                BeanCellMapping mapping = (BeanCellMapping) iterator.next();
                currentCellRef = new CellReference(mapping.getRow() + shift, mapping.getCol() );
                HSSFCell cell = getCell( cursor.getSheet(), mapping.getRow() + shift, mapping.getCol());
                Class type = mapping.getPropertyType( beans );
                Object value = getCellValue( cell, type );
                mapping.populateBean( value, beans );
            }
        } catch (Exception e) {
            throw new XLSDataReadException( currentCellRef==null?null:currentCellRef.toString(), "Can't read cell " + (currentCellRef==null?null:currentCellRef.toString()) + " on " + cursor.getSheetName() + " spreadsheet", e );
        }
        cursor.setCurrentRowNum( endRow + shift );
    }


    public SectionCheck getLoopBreakCondition() {
        return sectionCheck;
    }

    public void setLoopBreakCondition(SectionCheck sectionCheck) {
        this.sectionCheck = sectionCheck;
    }

    public void addBlockReader(XLSLoopBlockReader reader) {
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

    public List getMappings() {
        return beanCellMappings;
    }

    private Object getCellValue(HSSFCell cell, Class type) {
        if( cell == null ){
            return null;
        }
        Object value = null;
        if( type.getName().indexOf("java.util.Date") >=0 ){
            value = cell.getDateCellValue();
        }else{
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
