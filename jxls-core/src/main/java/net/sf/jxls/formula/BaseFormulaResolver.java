package net.sf.jxls.formula;

import org.apache.poi.hssf.util.CellReference;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.parser.Cell;

import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.util.Set;
import java.util.HashSet;
import java.util.List;

/**
 * Base class for {@link FormulaResolver} interface implementations
 * @author Leonid Vysochyn
 */
public abstract class BaseFormulaResolver implements FormulaResolver{
    protected static final String regexCellRef = "[a-zA-Z]+[0-9]+";
    protected static final Pattern regexCellRefPattern = Pattern.compile( regexCellRef );
    protected static final String regexCellCharPart = "[0-9]+";
    protected static final String regexCellDigitPart = "[a-zA-Z]+";
    protected String cellRangeSeparator = ":";
    protected static String formulaListRangeToken = "@";
    protected final Log log = LogFactory.getLog(getClass());

    Set findRefCells(String formulaString) {
        Set refCells = new HashSet();
        Matcher refCellMatcher = regexCellRefPattern.matcher( formulaString );
        while( refCellMatcher.find() ){
            refCells.add( refCellMatcher.group() );
        }
        return refCells;
    }

    String buildCommaSeparatedListOfCells(String refSheetName, List cells) {
        String listOfCells = "";
        for (int i = 0; i < cells.size() - 1; i++) {
            String cell = (String) cells.get(i);
            listOfCells += getRefCellName(refSheetName, cell) + ",";
        }
        listOfCells += getRefCellName( refSheetName, (String) cells.get( cells.size() - 1 ));
        return listOfCells;
    }

    String detectCellRange(String refSheetName, List cells) {
        String firstCell = (String) cells.get( 0 );
        String range = firstCell;
        if( firstCell != null && firstCell.length() > 0 ){
            if( isRowRange(cells) || isColumnRange(cells) ){
                String lastCell = (String) cells.get( cells.size() - 1 );
                range = getRefCellName(refSheetName, firstCell) + cellRangeSeparator + lastCell.toUpperCase();
            }else{
                range = buildCommaSeparatedListOfCells(refSheetName, cells );
            }
        }
        return range;
    }

    String getRefCellName(String refSheetName, String cellName){
        if( refSheetName == null ){
            return cellName.toUpperCase();
        }else{
            return refSheetName + "!" + cellName.toUpperCase();
        }
    }

    boolean isColumnRange(List cells) {
        String firstCell = (String) cells.get( 0 );
        boolean isColumnRange = true;
        if( firstCell != null && firstCell.length() > 0 ){
            String firstCellCharPart = firstCell.split(regexCellCharPart)[0];
            String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
            int cellNumber = Integer.parseInt( firstCellDigitPart );
            String nextCell, cellCharPart, cellDigitPart;
            for (int i = 1; i < cells.size() && isColumnRange; i++) {
                nextCell = (String) cells.get(i);
                cellCharPart = nextCell.split( regexCellCharPart )[0];
                cellDigitPart = nextCell.split( regexCellDigitPart )[1];
                if( !firstCellCharPart.equalsIgnoreCase( cellCharPart ) || Integer.parseInt(cellDigitPart) != ++cellNumber ){
                    isColumnRange = false;
                }
            }
        }
        return isColumnRange;
    }

    boolean isRowRange(List cells) {
        String firstCell = (String) cells.get( 0 );
        boolean isRowRange = true;
        if( firstCell != null && firstCell.length() > 0 ){
            String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
            String nextCell, cellDigitPart;
            CellReference cellRef = new CellReference( firstCell );
            int cellNumber = cellRef.getCol();
            for (int i = 1; i < cells.size() && isRowRange; i++) {
                nextCell = (String) cells.get(i);
                cellDigitPart = nextCell.split( regexCellDigitPart )[1];
                cellRef = new CellReference( nextCell );
                if( !firstCellDigitPart.equalsIgnoreCase( cellDigitPart ) || cellRef.getCol() != ++cellNumber ){
                    isRowRange = false;
                }
            }
        }
        return isRowRange;
    }

    /**
     * @return Formula string that should be set into Excel cell using POI
     */
    String replaceListRanges(Formula formula) {
        String codedFormula = formula.getFormula();
        Sheet sheet = formula.getSheet();
        String appliedFormula = "";
        String delimiter = formulaListRangeToken;
        int index = codedFormula.indexOf(delimiter);
        boolean isExpression = false;
        while (index >= 0) {
            String token = codedFormula.substring(0, index);
            if (isExpression) {
                // this is formula coded expression variable
                // look into the listRanges to see do we have cell range for it
                if (sheet.getListRanges().containsKey(token)) {
                    appliedFormula += ((ListRange) sheet.getListRanges().get(token)).toExcelCellRange();
                } else if (sheet.getNamedCells().containsKey(token)) {
                    appliedFormula += ((Cell) sheet.getNamedCells().get(token)).toCellName();
                } else {
                    log.warn("can't find list range or named cell for " + token);
                    // returning null if we don't have given list range or named cell so we don't need to set formula to avoid error
                    return null;
                }
            } else {
                appliedFormula += token;
            }
            codedFormula = codedFormula.substring(index + 1);
            index = codedFormula.indexOf(delimiter);
            isExpression = !isExpression;
        }
        appliedFormula += codedFormula;
        return appliedFormula;
    }

    public static String getFormulaListRangeToken() {
        return formulaListRangeToken;
    }

    public static void setFormulaListRangeToken(String formulaListRangeToken) {
        BaseFormulaResolver.formulaListRangeToken = formulaListRangeToken;
    }

}
