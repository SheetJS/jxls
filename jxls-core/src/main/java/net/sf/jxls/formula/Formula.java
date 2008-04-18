package net.sf.jxls.formula;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import net.sf.jxls.parser.Cell;
import net.sf.jxls.transformer.Sheet;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;


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
            FormulaPart formulaPart = (FormulaPart) f.formulaParts.get(i);
            formulaParts.add( new FormulaPart( formulaPart ) );
        }
        updateCellRefs();
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
        }
        return formula;
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

    private static final String regexFormulaPart = "[a-zA-Z]+[0-9]*\\([^@()]+\\)@[0-9]+";

    private static final Pattern regexFormulaPartPattern = Pattern.compile( regexFormulaPart );

    public String getActualFormula(){
        FormulaPart formulaPart;
        String actualFormula = "";
        for (Iterator iterator = formulaParts.iterator(); iterator.hasNext();) {
            formulaPart = (FormulaPart) iterator.next();
            actualFormula += formulaPart.getActualFormula();
        }
        return actualFormula;
    }

    public Set findRefCells() {
        Set refCells = new HashSet();
        for (Iterator iterator = formulaParts.iterator(); iterator.hasNext();) {
            FormulaPart formulaPart = (FormulaPart) iterator.next();
            refCells.addAll( formulaPart.getRefCells() );
        }
        return refCells;
    }

    public void parseFormula(){
        formulaParts.clear();
        Matcher formulaPartMatcher = regexFormulaPartPattern.matcher( formula );
        int end = 0;
        while( formulaPartMatcher.find() ){
            String formulaPartString = formula.substring(end, formulaPartMatcher.start());
            if( formulaPartString.length() > 0){
                formulaParts.add( new FormulaPart( formulaPartString, this) );
            }
            formulaParts.add( new FormulaPart(formulaPartMatcher.group(), this));
            end = formulaPartMatcher.end();
        }

        String endPart = formula.substring(end);
        if( endPart.length() > 0 ){
            formulaParts.add( new FormulaPart(endPart, this ));
        }
        updateCellRefs();
    }

    void updateCellRefs(){
        cellRefs = findRefCells();
    }

    public String toString() {
        return "Formula{" +
                "formula='" + formula + "'" +
                ", rowNum=" + rowNum +
                ", cellNum=" + cellNum +
                "}";
    }

    public boolean containsListRanges() {
        return formula.matches("[^)]*@.*");
    }

    public void removeCellRefs( Set cellRefsToRemove ){
        for (int i = 0; i < formulaParts.size(); i++) {
            FormulaPart formulaPart = (FormulaPart) formulaParts.get(i);
            formulaPart.removeCellRefs( cellRefsToRemove );
        }
        updateCellRefs();
    }

    public void updateReplacedRefCellsCollection(){
        for (Iterator iterator = formulaParts.iterator(); iterator.hasNext();) {
            FormulaPart formulaPart = (FormulaPart) iterator.next();
            formulaPart.updateReplacedRefCellsCollection( );
        }

    }

}
