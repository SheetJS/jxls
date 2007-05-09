package net.sf.jxls.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.tag.Block;
import net.sf.jxls.transformer.Row;
import net.sf.jxls.parser.CellParser;
import net.sf.jxls.formula.Formula;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class SheetHelper {



    public static List findFormulas(Sheet sheet){
        return findFormulas( sheet, new Block(null, 0, sheet.getHssfSheet().getLastRowNum() ) );
    }

    public static List findFormulas(Sheet sheet, Block block){
        List formulas = new ArrayList();
        for(int i = block.getStartRowNum(); i <= block.getEndRowNum(); i++){
            HSSFRow hssfRow = sheet.getHssfSheet().getRow( i );
            if( block.isRowBlock() ){
                formulas.addAll( findFormulasInRow(sheet, hssfRow) );
            }else{
                formulas.addAll( findFormulasInRow(sheet, hssfRow, block.getStartCellNum(), block.getEndCellNum() ));
            }
        }
        return formulas;
    }


    private static List findFormulasInRow(Sheet sheet, HSSFRow hssfRow, short startCellNum, short endCellNum) {
        List formulas = new ArrayList();
        if( hssfRow!=null ){
            Row row = new Row(sheet, hssfRow);
            short endNum = (short)Math.min( hssfRow.getLastCellNum(), endCellNum);
            for(short i = (short)Math.max(hssfRow.getFirstCellNum(), startCellNum); i <= endNum; i++){
                HSSFCell hssfCell = hssfRow.getCell( i );
                if( hssfCell!=null ){
                    CellParser cellParser = new CellParser(hssfCell, row, sheet.getConfiguration());
                    if( cellParser.parseCellFormula() != null && !cellParser.getCell().getFormula().isInline() ){
                        Formula formula = cellParser.getCell().getFormula();
                        formula.setSheet( sheet );
                        formulas.add( formula );
                    }
                }
            }
        }
        return formulas;
    }

    private static List findFormulasInRow(Sheet sheet, HSSFRow hssfRow) {
        List formulas = new ArrayList();
        if( hssfRow!=null ){
            Row row = new Row(sheet, hssfRow);
            CellParser cellParser;
            Formula formula;
            HSSFCell hssfCell;
            for(short i = hssfRow.getFirstCellNum(); i <= hssfRow.getLastCellNum(); i++){
                hssfCell = hssfRow.getCell( i );
                if( hssfCell!=null ){
                    cellParser = new CellParser(hssfCell, row, sheet.getConfiguration());
                    if( cellParser.parseCellFormula() != null && !cellParser.getCell().getFormula().isInline() ){
                        formula = cellParser.getCell().getFormula();
                        formula.setSheet( sheet );
                        formulas.add( formula );
                    }
                }
            }
        }
        return formulas;
    }
}
