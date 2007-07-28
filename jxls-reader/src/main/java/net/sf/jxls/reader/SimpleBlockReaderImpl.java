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
public class SimpleBlockReaderImpl extends BaseBlockReader implements SimpleBlockReader{
    protected final Log log = LogFactory.getLog(getClass());

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
        final int currentRowNum = cursor.getCurrentRowNum();
        final int rowShift = currentRowNum - startRow;
        BeanCellMapping mapping;
            for (Iterator iterator = beanCellMappings.iterator(); iterator.hasNext();) {
                mapping = (BeanCellMapping) iterator.next();
//                Object value;
                try {
//                    value = readCellValue(cursor.getSheet(), mapping.getRow() + rowShift, mapping.getCol(), mapping.getPropertyType( beans ));
                    String dataString = readCellString(cursor.getSheet(), mapping.getRow() + rowShift, mapping.getCol() );
                    mapping.populateBean( dataString, beans );
                } catch (Exception e) {
                    String message = "Can't read cell " + getCellName( mapping, rowShift ) + " on " + cursor.getSheetName() + " spreadsheet";
                    if( Configuration.getInstance().isSkipErrors() )    {
                        log.error( message, e);
                    }else{
                        throw new XLSDataReadException(getCellName( mapping, rowShift ), "Can't read cell " + getCellName( mapping, rowShift ) + " on " + cursor.getSheetName() + " spreadsheet", e );
                    }
                }
            }
        cursor.setCurrentRowNum( endRow + rowShift );
    }

    private String readCellString(HSSFSheet sheet, int rowNum, short cellNum) {
        HSSFCell cell = getCell(sheet, rowNum, cellNum);
        return getCellString( cell );
    }

    private String getCellString(HSSFCell cell) {
        String dataString = null;
        if( cell != null ){
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    dataString = cell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    double value = cell.getNumericCellValue();
                    if( ((int)value) == value ){
                        dataString = Integer.toString( (int)value );
                    }else{
                        dataString = Double.toString( cell.getNumericCellValue());
                    }
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    dataString = Boolean.toString( cell.getBooleanCellValue() );
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
        return dataString;
    }

    private Object readCellValue(HSSFSheet sheet, int rowNum, short cellNum, Class type) {
        HSSFCell cell = getCell(sheet, rowNum, cellNum);
        return getCellValue( cell, type );
    }

    private String getCellName( BeanCellMapping mapping, int rowShift ){
        CellReference currentCellRef = new CellReference(mapping.getRow() + rowShift, mapping.getCol());
        return currentCellRef.toString();
    }


    public SectionCheck getLoopBreakCondition() {
        return sectionCheck;
    }

    public void setLoopBreakCondition(SectionCheck sectionCheck) {
        this.sectionCheck = sectionCheck;
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
        if(isDate(type)){
            value = cell.getDateCellValue();
        }else{
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    value = new Double(cell.getNumericCellValue());
                    // if we expect String type instead of numeric then we converting Double value to String
                    if( !type.isInstance(value) ){
                        double doubleValue = ((Double)value).doubleValue();
                        if( Math.floor( doubleValue ) == doubleValue ){
                            value = Long.toString(((Double) value).longValue());
                        }else{
                            value = value.toString();
                        }
                    }
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

    private boolean isDate(Class type) {
        return type.getName().indexOf("java.util.Date") >=0;
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
