package net.sf.jxls.sample;

import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.parser.Cell;
import net.sf.jxls.parser.Expression;
import net.sf.jxls.parser.Property;
import net.sf.jxls.processor.CellProcessor;
import net.sf.jxls.sample.model.Employee;
import net.sf.jxls.util.Util;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * WARNING! This class does not work since 0.8.7 jXLS release as there were changes in expression parsing introduced
 * for full JEXL support.
 * If you need to apply conditional formatting to some cells or row you may use jx:if tag.
 * @deprecated
 * @author Leonid Vysochyn
 */
public class RedCellProcessor2  implements CellProcessor {

    static CellStyle hssfCellStyle;
    String beanName;

    Map rowStyles = new HashMap();

    public RedCellProcessor2(String collectionName) {
        this.beanName = collectionName.replace('.', '_');
    }

    public void processCell(final Cell cell, final Map namedCells) {
        if( cell.getExpressions().size()>0 ){
            Expression expression = (Expression) cell.getExpressions().get(0);
            Property property = (Property)expression.getProperties().get(0);
            if (property != null && property.getBeanName() != null && property.getBeanName().indexOf(beanName) >= 0 && property.getBean() instanceof Employee) {
                Employee employee = (Employee) property.getBean();
                if (employee.getPayment().doubleValue() >= 2000) {
                        org.apache.poi.ss.usermodel.Cell hssfCell = cell.getPoiCell();
                        CellStyle newStyle = duplicateStyle( cell, property.getPropertyNameAfterLastDot() );
                        newStyle.setFillForegroundColor( HSSFColor.RED.index );
                        newStyle.setFillPattern( CellStyle.SOLID_FOREGROUND );
                        hssfCell.setCellStyle( newStyle );
                }
            }
        }
    }

    CellStyle duplicateStyle( Cell cell, String key ){
        if( rowStyles.containsKey( key ) ){
            return (CellStyle) rowStyles.get( key );
        }
        CellStyle newStyle =  Util.duplicateStyle( cell.getRow().getSheet().getPoiWorkbook(), cell.getPoiCell().getCellStyle() );
        rowStyles.put( key, newStyle );
        return newStyle;
    }

}
