package net.sf.jxls.reader;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

/**
 * @author Leonid Vysochyn
 */
public class OffsetCellCheckImpl implements OffsetCellCheck {
    Object value;
    short offset;

    public OffsetCellCheckImpl() {
    }

    public OffsetCellCheckImpl(short offset, Object value) {
        this.offset = offset;
        this.value = value;
    }

    public OffsetCellCheckImpl(Object value) {
        this.value = value;
    }


    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public short getOffset() {
        return offset;
    }

    public void setOffset(short offset) {
        this.offset = offset;
    }

    public boolean isCheckSuccessful(HSSFCell cell) {
        Object obj = getCellValue( cell, value );
        if( value == null ){
            return obj == null;
        }else{
            return value.equals( obj );
        }
    }

    public boolean isCheckSuccessful(HSSFRow row) {
        if( row == null ){
            return value == null;
        }else{
            HSSFCell cell = row.getCell( offset );
            return isCheckSuccessful( cell );
        }
    }

    private Object getCellValue(HSSFCell cell, Object obj) {
        Object value = null;
        if (obj instanceof String) {
            value = readStringValue(cell);
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
            c.setTime( cell.getDateCellValue() );
            value = c;
        } else if (obj instanceof Boolean){
            if( cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN ){
                value = (cell.getBooleanCellValue()) ? Boolean.TRUE : Boolean.FALSE;
            }else if( cell.getCellType() == HSSFCell.CELL_TYPE_STRING ){
                value = Boolean.valueOf( cell.getStringCellValue() );
            }else{
                value = Boolean.FALSE;
            }
        }
        return value;
    }

    private String readStringValue(HSSFCell cell) {
        String value = null;
        int cellType= cell==null ? HSSFCell.CELL_TYPE_BLANK : cell.getCellType();
        switch (cellType) {
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue() != null ? cell.getStringCellValue().trim() : null;
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                value = Double.toString(cell.getNumericCellValue()).trim();
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                value = "";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
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
}
