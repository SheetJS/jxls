package net.sf.jxls.reader;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;

/**
 * @author Leonid Vysochyn
 */
public class OffsetRowCheckImpl implements OffsetRowCheck {

    List cellChecks = new ArrayList();
    int offset;


    public OffsetRowCheckImpl() {
    }

    public OffsetRowCheckImpl(int offset) {
        this.offset = offset;
    }

    public OffsetRowCheckImpl(List cellChecks) {
        this.cellChecks = cellChecks;
    }

    public int getOffset() {
        return offset;
    }

    public void setOffset(int offset) {
        this.offset = offset;
    }

    public List getCellChecks() {
        return cellChecks;
    }

    public void setCellChecks(List cellChecks) {
        this.cellChecks = cellChecks;
    }

    public boolean isCheckSuccessful(HSSFRow row) {
        if( cellChecks.isEmpty() ){
            return isRowEmpty( row );
        }
        for (int i = 0; i < cellChecks.size(); i++) {
            OffsetCellCheck offsetCellCheck = (OffsetCellCheck) cellChecks.get(i);
            if( !offsetCellCheck.isCheckSuccessful( row ) ){
                return false;
            }
        }
        return true;
    }

    public boolean isCheckSuccessful(XLSRowCursor cursor) {
        if( !cursor.hasNext() ){
            return isCellChecksEmpty();
        }
        HSSFRow row = cursor.getSheet().getRow( offset + cursor.getCurrentRowNum() );
        if( row == null ){
            return cellChecks.isEmpty();
        }else{
            return isCheckSuccessful( row );
        }
    }

    private boolean isCellChecksEmpty() {
        if( cellChecks.isEmpty() ){
            return true;
        }else{
            for (int i = 0; i < cellChecks.size(); i++) {
                OffsetCellCheck offsetCellCheck = (OffsetCellCheck) cellChecks.get(i);
                if( !isCellCheckEmpty(offsetCellCheck) ){
                    return false;
                }
            }
            return true;
        }
    }

    private boolean isCellCheckEmpty(OffsetCellCheck cellCheck) {
        if( cellCheck.getValue() == null ){
            return true;
        }
        if( cellCheck.getValue().toString().trim().equals("") ){
            return true;
        }else{
            return false;
        }
    }


    public void addCellCheck(OffsetCellCheck cellCheck) {
        cellChecks.add( cellCheck );
    }

    private boolean isRowEmpty(HSSFRow row) {
        if( row == null ){
            return true;
        }
        for(short i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++){
            HSSFCell cell = row.getCell( i );
            if( !isCellEmpty( cell ) ){
                return false;
            }
        }
        return true;
    }

    private boolean isCellEmpty(HSSFCell cell) {
        if( cell == null ){
            return true;
        }else{
            switch( cell.getCellType() ){
                case HSSFCell.CELL_TYPE_BLANK:
                    return true;
                case HSSFCell.CELL_TYPE_STRING:
                    String cellValue = cell.getStringCellValue();
                    return cellValue == null || cellValue.length() == 0 || cellValue.trim().length() == 0;
                default:
                    return false;
            }
        }
    }
}
