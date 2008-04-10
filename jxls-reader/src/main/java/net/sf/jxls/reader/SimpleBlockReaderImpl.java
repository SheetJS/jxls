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

    static{
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

    public XLSReadStatus read(XLSRowCursor cursor, Map beans)  {
        readStatus.clear();
        final int currentRowNum = cursor.getCurrentRowNum();
        final int rowShift = currentRowNum - startRow;
        BeanCellMapping mapping;
            for (Iterator iterator = beanCellMappings.iterator(); iterator.hasNext();) {
                mapping = (BeanCellMapping) iterator.next();
                try {
                    String dataString = readCellString(cursor.getSheet(), mapping.getRow() + rowShift, mapping.getCol() );
                    mapping.populateBean( dataString, beans );
                } catch (Exception e) {
                    String message = "Can't read cell " + getCellName( mapping, rowShift ) + " on " + cursor.getSheetName() + " spreadsheet";
                    readStatus.addMessage( new XLSReadMessage(message, e));
                    if( ReaderConfig.getInstance().isSkipErrors() )    {
                        if( log.isWarnEnabled() ){
                            log.warn( message );
                        }
                    }else{
                        readStatus.setStatusOK( false );
                        throw new XLSDataReadException(getCellName( mapping, rowShift ), "Can't read cell " + getCellName( mapping, rowShift ) + " on " + cursor.getSheetName() + " spreadsheet", readStatus);
                    }
                }
            }
        cursor.setCurrentRowNum( endRow + rowShift );
        return readStatus;
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
                    dataString = readNumericCell(cell);
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    dataString = Boolean.toString( cell.getBooleanCellValue() );
                    break;
                case HSSFCell.CELL_TYPE_BLANK:
                    break;
                case HSSFCell.CELL_TYPE_ERROR:
                    break;
                case HSSFCell.CELL_TYPE_FORMULA:
                    // attempt to read formula cell as numeric cell
                    dataString = readNumericCell(cell);
                    break;
                default:
                    break;
            }
        }
        return dataString;
    }

    private String readNumericCell(HSSFCell cell) {
        double value;
        String dataString;
        value = cell.getNumericCellValue();
        if( ((int)value) == value ){
            dataString = Integer.toString( (int)value );
        }else{
            dataString = Double.toString( cell.getNumericCellValue());
        }
        return dataString;
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

    private HSSFCell getCell(HSSFSheet sheet, int rowNum, short cellNum){
        HSSFRow row = sheet.getRow( rowNum );
        if( row == null ){
            return null;
        }else{
            return row.getCell( cellNum );
        }
    }

}
