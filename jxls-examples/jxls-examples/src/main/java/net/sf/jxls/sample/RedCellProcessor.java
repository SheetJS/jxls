package net.sf.jxls.sample;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.processor.CellProcessor;
import net.sf.jxls.parser.Property;
import net.sf.jxls.parser.Expression;
import net.sf.jxls.sample.model.Employee;

import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class RedCellProcessor implements CellProcessor {

    static HSSFCellStyle hssfCellStyle;
    public static final String RED_CELL = "red";
    String beanName;

    public RedCellProcessor(String collectionName) {
        this.beanName = collectionName.replace('.', '_');
    }

    public void processCell(final Cell cell, final Map namedCells) {
        if( cell.getExpressions().size()>0 ){
            Expression expression = (Expression) cell.getExpressions().get(0);
            Property property = (Property)expression.getProperties().get(0);
            if (property != null && property.getBeanName() != null && property.getBeanName().indexOf(beanName) >= 0 && property.getBean() instanceof Employee) {
                Employee employee = (Employee) property.getBean();
                if (employee.getPayment().doubleValue() >= 2000) {
                    if (namedCells.containsKey(RED_CELL + "_" + property.getPropertyNameAfterLastDot())) {
                        Cell redCell = (Cell) namedCells.get(RED_CELL + "_" + property.getPropertyNameAfterLastDot());
                        HSSFCellStyle redStyle = redCell.getHssfCell().getCellStyle();
                        cell.getHssfCell().setCellStyle(redStyle);
                    }
                }
            }
        }
    }
}
