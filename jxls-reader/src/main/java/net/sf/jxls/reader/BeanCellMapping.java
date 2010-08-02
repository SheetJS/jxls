package net.sf.jxls.reader;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.util.CellReference;

/**
 * @author Leonid Vysochyn
 */
public class BeanCellMapping {
    protected final Log log = LogFactory.getLog(getClass());

    int row;
    short col;
    String propertyName;
    String beanKey;
    String cell;
    String type = null;
    boolean nullAllowed = false;

    static {
        ReaderConfig.getInstance();
    }

    public BeanCellMapping(int rowNum, short cellNum, String fullPropertyName) {
        this.row = rowNum;
        this.col = cellNum;
        this.beanKey = extractBeanName(fullPropertyName);
        this.propertyName = extractPropertyName(fullPropertyName);
    }

    public BeanCellMapping(String cell, String fullPropertyName) {
        setCell(cell);
        this.beanKey = extractBeanName(fullPropertyName);
        this.propertyName = extractPropertyName(fullPropertyName);
    }

    public BeanCellMapping(int rowNum, short cellNum, String beanKey, String propertyName) {
        this.row = rowNum;
        this.col = cellNum;
        this.beanKey = beanKey;
        this.propertyName = propertyName;
    }

    public BeanCellMapping(String cell, String beanKey, String propertyName) {
        setCell(cell);
        this.beanKey = beanKey;
        this.propertyName = propertyName;
    }

    public BeanCellMapping() {
    }


    public String getBeanKey() {
        return beanKey;
    }

    public void setBeanKey(String beanKey) {
        this.beanKey = beanKey;
    }

    public String getFullPropertyName() {
        return beanKey + "." + propertyName;
    }

    public void setFullPropertyName(String fullPropertyName) {
        this.beanKey = extractBeanName(fullPropertyName);
        this.propertyName = extractPropertyName(fullPropertyName);
    }

    private String extractPropertyName(String fullPropertyName) {
        if (fullPropertyName == null) {
            return null;
        }
        int dotIndex = fullPropertyName.indexOf('.');
        if (dotIndex < 0) {
            throw new IllegalArgumentException("Full property name must contain period. Can't extract bean property name from " + fullPropertyName);
        } else {
            return fullPropertyName.substring(dotIndex + 1);
        }
    }

    private String extractBeanName(String fullPropertyName) {
        if (fullPropertyName == null) {
            return null;
        }
        int dotIndex = fullPropertyName.indexOf('.');
        if (dotIndex < 0) {
            throw new IllegalArgumentException("Full property name must contain period. Can't extract bean name from " + fullPropertyName);
        } else {
            return fullPropertyName.substring(0, dotIndex);
        }
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public short getCol() {
        return col;
    }

    public void setCol(short col) {
        this.col = col;
    }


    public String getCell() {
        return cell;
    }

    public void setCell(String cell) {
        this.cell = cell;
        CellReference cellRef = new CellReference(cell);
        row = cellRef.getRow();
        col = cellRef.getCol();
    }

    public String getPropertyName() {
        return propertyName;
    }

    public void setPropertyName(String propertyName) {
        this.propertyName = propertyName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public boolean isNullAllowed() {
        return nullAllowed;
    }

    public void setNullAllowed(boolean nullAllowed) {
        this.nullAllowed = nullAllowed;
    }

    public void populateBean(String dataString, Map beans) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, ClassNotFoundException {
        Object bean;
        if (beans.containsKey(beanKey)) {
            bean = beans.get(beanKey);
            Class dataType = getPropertyType(beans);

            Object value = null;
            if (!(nullAllowed && dataString == null)) { // set only if null is not allowed!
                value = ConvertUtils.convert(dataString, dataType);
            }

            PropertyUtils.setProperty(bean, propertyName, value);
        } else {
            if (log.isWarnEnabled()) {
                log.warn("Can't find bean under the key=" + beanKey);
            }
        }
    }

    // Added feature to set type by type-attribute in XML
    public Class getPropertyType(Map beans) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, ClassNotFoundException {
        Object bean;
        if (beans.containsKey(beanKey)) {
            bean = beans.get(beanKey);
            // ZJ
            if (type == null) {
                return PropertyUtils.getPropertyType(bean, propertyName);
            } else {
                return Class.forName(type);
            }
        }
        return Object.class;
    }

    public String getCellName() {
        CellReference cellRef = new CellReference(row, col, false, false);
        return cellRef.formatAsString();
    }

    public String toString() {
        return beanKey + ":" + propertyName;
    }
}
