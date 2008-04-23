package net.sf.jxls.tag;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

import net.sf.jxls.parser.Expression;
import net.sf.jxls.transformation.ResultTransformation;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.transformer.SheetTransformer;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class OutTag extends BaseTag {
    
    protected final Log log = LogFactory.getLog(getClass());
    
    public static final String TAG_NAME = "out";
    
    private Configuration configuration = new Configuration();
    private TagContext tagContext;
    private String expr;
    private String formula;
    private String label;

    public String getExpr() {
        return expr;
    }

    public void setExpr(String expr) {
        this.expr = expr;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public String getName() {
        return TAG_NAME;
    }

    public TagContext getTagContext() {
        return tagContext;
    }

    public void init(TagContext context) {
        this.tagContext = context;
    }

    public ResultTransformation process(SheetTransformer sheetTransformer) {
        
        ResultTransformation resultTransformation = new ResultTransformation(0);

        if (expr != null) {
            
            // process expression cell

            try {
                Block block = getTagContext().getTagBody();
                int rowNum = block.getStartRowNum();
                int cellNum = block.getStartCellNum();
                
                Sheet jxlsSheet = getTagContext().getSheet();
                if (jxlsSheet != null) {
                    HSSFSheet sheet = jxlsSheet.getHssfSheet();
                    if (sheet != null) {
                        HSSFRow row = sheet.getRow(rowNum);
                        if (row != null) {
                            HSSFCell cell = row.getCell((short) cellNum);
                            if (cell != null) {
                                
                                Object value = new Expression(expr, tagContext.getBeans(), configuration).evaluate();
                                if (value == null) {
                                    cell.setCellValue(new HSSFRichTextString(""));
                                } else if (value instanceof Double) {
                                    cell.setCellValue(((Double) value).doubleValue());
                                } else if (value instanceof BigDecimal) {
                                    cell.setCellValue(((BigDecimal) value).doubleValue());
                                } else if (value instanceof Date) {
                                    cell.setCellValue((Date) value);
                                }else if (value instanceof Calendar) {
                                    cell.setCellValue((Calendar) value);
                                } else if (value instanceof Integer) {
                                    cell.setCellValue(((Integer) value).intValue());
                                }else if (value instanceof Long) {
                                    cell.setCellValue(((Long) value).longValue());
                                } else {
                                    // fixing possible CR/LF problem
                                    String fixedValue = value.toString();
                                    if (fixedValue != null) {
                                        fixedValue = fixedValue.replaceAll("\r\n", "\n");
                                    }
                                    cell.setCellValue(new HSSFRichTextString(fixedValue));
                                }
                            }
                        }
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                log.error("Cell expression evaluation has failed.", e);
            }
        }
        
        return resultTransformation;
    }
}