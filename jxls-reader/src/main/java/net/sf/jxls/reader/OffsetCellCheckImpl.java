package net.sf.jxls.reader;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

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

    public boolean isCheckSuccessful(Cell cell) {
        Object obj = getCellValue(cell, value);
        if (value == null || value.toString().trim().length()==0) {
            return obj == null || obj.toString().trim().length()==0;
        } else {
            return value.equals(obj);
        }
    }

    public boolean isCheckSuccessful(Row row) {
        if (row == null) {
            return value == null;
        } else {
            Cell cell = row.getCell(offset);
            return isCheckSuccessful(cell);
        }
    }

    private Object getCellValue(Cell cell, Object obj) {
        if(cell == null ){
            return null;
        }
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

    private String readStringValue(Cell cell) {
        String value = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                value = Double.toString(cell.getNumericCellValue()).trim();
                break;
            case Cell.CELL_TYPE_BLANK:
                value = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                break;
            case Cell.CELL_TYPE_ERROR:
                break;
            case Cell.CELL_TYPE_FORMULA:
                break;
            default:
                break;
        }
        return value;
    }
}
