package net.sf.jxls.util;

import java.util.List;
import java.util.Map;

import net.sf.jxls.formula.Formula;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * @author Leonid Vysochyn
 */
public class FormulaUtil {

    /**
     * This method updates formula cells
     *
     * @param sheet            Sheet to update
     * @param ignoreUnresolved Flag indicating should unresolved formulas be removed or just ignored
     */
    public static void updateFormulas(HSSFSheet sheet, List formulas, Map listRanges, Map namedCells, boolean ignoreUnresolved) {
        for (int i = 0; i < formulas.size(); i++) {
            Formula formula = (Formula) formulas.get(i);
            String formulaString = formula.getAppliedFormula( listRanges, namedCells );
            HSSFRow hssfRow = sheet.getRow(formula.getRowNum().intValue());
            HSSFCell hssfCell = hssfRow.getCell(formula.getCellNum().shortValue());
            if (formulaString != null) {
                hssfCell.setCellFormula(formulaString);
            } else {
                if (!ignoreUnresolved) {
                    hssfCell.setCellValue(new HSSFRichTextString(""));
                    formulas.remove(i--);
                }
            }
        }
    }

    /**
     * @param formulas {@link List} of {@link Formula} objects
     * @param formula {@link Formula} object to check
     * @return true if given {@link Formula} already exists in formulas list
     */
    static boolean formulaExists(List formulas, Formula formula) {
        for (int i = 0; i < formulas.size(); i++) {
            Formula cur = (Formula) formulas.get(i);
            if (cur.getFormula().equals(formula.getFormula())) {
                return true;
            }
        }
        return false;
    }


}
