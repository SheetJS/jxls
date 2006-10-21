package net.sf.jxls.formula;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.controller.SheetCellFinder;

import java.util.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;


/**
 * Represents formula cell
 * @author Leonid Vysochyn
 */
public class Formula {
    protected final Log log = LogFactory.getLog(getClass());

    private String formula;
    private Integer rowNum;
    private Integer cellNum;
    static final String inlineFormulaToken = "#";
    static final String formulaListRangeToken = "@";

    private Sheet sheet;

    private Set cellRefs = new HashSet();

    private List formulaParts = new ArrayList();

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Formula(String formula) {
        this.formula = formula;
        parseFormula();
    }

    public Formula() {
    }

    public Formula(Formula f){
        this.formula = f.formula;
        this.sheet = f.getSheet();
        for (int i = 0; i < f.formulaParts.size(); i++) {
            Object formulaPart = f.formulaParts.get(i);
            if( formulaPart instanceof String ){
                formulaParts.add(formulaPart.toString());
            }else if(formulaPart instanceof CellRef){
                CellRef cellRef = new CellRef( formulaPart.toString(), this );
                formulaParts.add( cellRef );
                cellRefs.add( cellRef );
            }
        }
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public Integer getRowNum() {
        return rowNum;
    }

    public void setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
    }

    public Integer getCellNum() {
        return cellNum;
    }

    public void setCellNum(Integer cellNum) {
        this.cellNum = cellNum;
    }

    public Set getCellRefs() {
        return cellRefs;
    }

    public List getFormulaParts() {
        return formulaParts;
    }

    public boolean isInline() {
        return formula.indexOf(inlineFormulaToken) >= 0;
    }

    public String getInlineFormula(int n) {
        if (isInline()) {
            return formula.replaceAll(inlineFormulaToken, Integer.toString(n));
        } else {
            return formula;
        }
    }

    /**
     * @return Formula string that should be set into Excel cell using POI
     */
    public String getAppliedFormula(Map listRanges, Map namedCells) {
        String codedFormula = formula;
        String appliedFormula = "";
        String delimiter = formulaListRangeToken;
        int index = codedFormula.indexOf(delimiter);
        boolean isExpression = false;
        while (index >= 0) {
            String token = codedFormula.substring(0, index);
            if (isExpression) {
                // this is formula coded expression variable
                // look into the listRanges to see do we have cell range for it
                if (listRanges.containsKey(token)) {
                    appliedFormula += ((ListRange) listRanges.get(token)).toExcelCellRange();
                } else if (namedCells.containsKey(token)) {
                    appliedFormula += ((Cell) namedCells.get(token)).toCellName();
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

    String adjust(SheetCellFinder cellFinder){
//        String adjustedFormula = formula;
//        Set refCells = findRefCells();
//        for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
//            String refCell = (String) iterator.next();
//            String newCell = cellFinder.findCell( refCell );
//            adjustedFormula = adjustedFormula.replaceAll( refCell, newCell );
//        }
//        formula = adjustedFormula;
        return formula;
    }

    private static final String regexCellRef = "([a-zA-Z]+[a-zA-Z0-9]*![a-zA-Z]+[0-9]+|[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final Pattern regexCellRefPattern = Pattern.compile( regexCellRef );

    public String getActualFormula(){
        Object formulaPart;
        String actualFormula = "";
        for (Iterator iterator = formulaParts.iterator(); iterator.hasNext();) {
            formulaPart =  iterator.next();
            actualFormula += formulaPart.toString();
        }
        return actualFormula;
    }

    public Set findRefCells() {
        Set refCells = new HashSet();
        Matcher refCellMatcher = regexCellRefPattern.matcher( formula );
        while( refCellMatcher.find() ){
            refCells.add( refCellMatcher.group() );
        }
        return refCells;
    }

    public void parseFormula(){
        formulaParts.clear();
        cellRefs.clear();
        Matcher refCellMatcher = regexCellRefPattern.matcher( formula );
        int end = 0;
        CellRef cellRef;
        while( refCellMatcher.find() ){
            formulaParts.add( formula.substring( end, refCellMatcher.start() ) );
            cellRef = new CellRef( refCellMatcher.group(), this );
            formulaParts.add( cellRef );
            cellRefs.add( cellRef );
            end = refCellMatcher.end();
        }
        formulaParts.add( formula.substring( end ));
    }

    public String toString() {
        return "Formula{" +
                "formula='" + formula + "'" +
                ", rowNum=" + rowNum +
                ", cellNum=" + cellNum +
                "}";
    }

    public boolean containsListRanges() {
        return formula.indexOf( formulaListRangeToken ) >= 0;
    }

    public void replaceCellRef(CellRef cellRef, List rangeFormulaParts) {
        for (int i = 0; i < formulaParts.size(); i++) {
            Object formulaPart = formulaParts.get(i);
            if( formulaPart == cellRef ){
                replaceFormulaPart( i, rangeFormulaParts );
                replaceCellRefs( cellRef, rangeFormulaParts );
                break;
            }
        }
    }

    public void removeCellRefs( Set cellRefsToRemove ){
        List formulaPartIndexesToRemove = new ArrayList();
        Object prevFormulaPart = null;
        Object nextFormulaPart = null;
        for (int i = 0; i < formulaParts.size(); i++) {
            Object formulaPart = formulaParts.get(i);
            if( cellRefsToRemove.contains( formulaPart ) ){
                formulaPartIndexesToRemove.add( new Integer( i ) );
                if( i > 0 ){
                    prevFormulaPart = formulaParts.get( i - 1 );
                }
                if( i < formulaParts.size() - 1 ){
                    nextFormulaPart = formulaParts.get( i + 1 );
                }else{
                    nextFormulaPart = null;
                }
                if( prevFormulaPart != null ){
                    if( prevFormulaPart.toString().equals(",") ){
                        formulaPartIndexesToRemove.add( new Integer(i - 1) );
                    }else if( nextFormulaPart != null && nextFormulaPart.toString().equals( "," )){
                        formulaPartIndexesToRemove.add( new Integer(i + 1) );
                    }
                }
            }
        }
        int shift = 0;
        for (int i = 0; i < formulaPartIndexesToRemove.size(); i++) {
            int index =  ((Integer) formulaPartIndexesToRemove.get(i)).intValue() ;
            formulaParts.remove( index - shift );
            shift++;
        }
        cellRefs.removeAll( cellRefsToRemove );
    }

    private void replaceCellRefs(CellRef cellRef, List rangeFormulaParts) {
        cellRefsToRemove.add( cellRef );
        for (int i = 0; i < rangeFormulaParts.size(); i++) {
            Object formulaPart = rangeFormulaParts.get(i);
            if( formulaPart instanceof CellRef ){
                cellRefsToAdd.add( formulaPart );
            }
        }
    }

    List cellRefsToRemove = new ArrayList();
    List cellRefsToAdd = new ArrayList();
    public void updateReplacedRefCellsCollection(){
        CellRef cellRef;
        for (int i = 0, size = cellRefsToRemove.size(); i < size; i++) {
            cellRef = (CellRef) cellRefsToRemove.get(i);
            cellRefs.remove( cellRef );
        }
        cellRefsToRemove.clear();
        Object cellRef2;
        for (int i = 0, size = cellRefsToAdd.size(); i < size; i++) {
            cellRef2 = cellRefsToAdd.get(i);
            cellRefs.add( cellRef2 );
        }
        cellRefsToAdd.clear();
    }

    private void replaceFormulaPart(int pos, List rangeFormulaParts) {
        formulaParts.remove( pos );
        formulaParts.addAll( pos, rangeFormulaParts );
    }
}
