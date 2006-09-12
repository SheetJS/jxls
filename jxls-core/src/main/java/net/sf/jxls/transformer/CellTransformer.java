package net.sf.jxls.transformer;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.parser.Expression;
import net.sf.jxls.formula.Formula;

import java.math.BigDecimal;
import java.util.Date;

/**
 * Cell transformation class
 * @author Leonid Vysochyn
 */
public class CellTransformer {
    protected final Log log = LogFactory.getLog(getClass());


    private Configuration configuration;

    public CellTransformer(Configuration configuration) {
        if( configuration!=null ){
            this.configuration = configuration;
        }else{
            this.configuration = new Configuration();
        }
    }

    void transform( Cell cell ){
        try {
            if (cell.getHssfCell() != null && cell.getHssfCell().getCellType() == HSSFCell.CELL_TYPE_STRING) {
                if (cell.getCollectionProperty() == null) {
                    if (cell.getFormula() == null) {
                            if( cell.getExpressions().size() == 0 ){
                                if( cell.getMetaInfo() !=null ){
                                    cell.getHssfCell().setEncoding( HSSFCell.ENCODING_UTF_16 );
                                    cell.getHssfCell().setCellValue( cell.getStringCellValue() );
                                }
                            }else if (cell.getExpressions().size() == 1) {
                                Object value = ((Expression) cell.getExpressions().get(0)).evaluate();
                                if (value == null) {
                                    cell.getHssfCell().setCellValue("");
                                } else if (value instanceof Double) {
                                    cell.getHssfCell().setCellValue(((Double) value).doubleValue());
                                } else if (value instanceof BigDecimal) {
                                    cell.getHssfCell().setCellValue(((BigDecimal) value).doubleValue());
                                } else if (value instanceof Date) {
                                    cell.getHssfCell().setCellValue((Date) value);
                                } else if (value instanceof Integer) {
                                    cell.getHssfCell().setCellValue(((Integer) value).intValue());
                                }else if (value instanceof Long) {
                                    cell.getHssfCell().setCellValue(((Long) value).longValue());
                                } else {
                                    // fixing possible CR/LF problem
                                    String fixedValue = value.toString();
                                    if (fixedValue != null) {
                                        fixedValue = fixedValue.replaceAll("\r\n", "\n");
                                    }
                                    cell.getHssfCell().setEncoding( HSSFCell.ENCODING_UTF_16 );
                                    cell.getHssfCell().setCellValue(fixedValue);
                                }
                            } else {
                                if (cell.getExpressions().size() > 1) {
                                    String value = "";
                                    for (int i = 0; i < cell.getExpressions().size(); i++) {
                                        Expression expr = (Expression) cell.getExpressions().get(i);
                                        Object propValue = expr.evaluate();
                                        if (propValue != null) {
                                            value += propValue.toString();
                                        }
                                    }
                                    cell.getHssfCell().setEncoding( HSSFCell.ENCODING_UTF_16 );
                                    cell.getHssfCell().setCellValue(value);
                                }
                            }
                    }
                    else {
                        processFormulaCell( cell );
                    }
                } else {
                    String value = "";
                    for (int i = 0; i < cell.getExpressions().size(); i++) {
                        Expression expr = (Expression) cell.getExpressions().get(i);
                        if (expr.getCollectionProperty() == null) {
                            value += expr.evaluate();
                        } else {
                            value += configuration.getStartExpressionToken() + expr.getExpression() + configuration.getEndExpressionToken();
                        }
                    }
                    cell.getHssfCell().setEncoding( HSSFCell.ENCODING_UTF_16 );
                    cell.getHssfCell().setCellValue(value);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.debug("Can't parse expression");
        }
    }

    private static void processFormulaCell(Cell cell) {
        // processing formula
        Formula formula = cell.getFormula();
        if (formula.isInline()) {
            if (cell.getCollectionName() != null) {
        // simple copy of inline formula template
        // it will be processed when individual rows are processed
                cell.getHssfCell().setCellValue( cell.getStringCellValue() );
            } else {
        // processing of inline formulaString template
                String formulaString = formula.getInlineFormula(cell.getRow().getHssfRow().getRowNum() + 1);
                cell.getHssfCell().setCellFormula(formulaString);
            }
        }
    }
}
