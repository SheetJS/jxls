package net.sf.jxls.formula;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * Represents formula part 
 * @author Leonid Vysochyn
 */
public class FormulaPart {
    protected final Log log = LogFactory.getLog(getClass());

    public static char defaultValueToken = '@';

    Formula parentFormula;
    String formulaPartString;
    List    parts = new ArrayList();
    List cellRefs = new ArrayList();
    List cellRefsToRemove = new ArrayList();
    List cellRefsToAdd = new ArrayList();


    Integer defaultValue = null;

    private static final String regexCellRef = "([a-zA-Z]+[a-zA-Z0-9]*![a-zA-Z]+[0-9]+|[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final Pattern regexCellRefPattern = Pattern.compile( regexCellRef );


    public FormulaPart(String formulaPartString, Formula parentFormula) {
        this.formulaPartString = formulaPartString;
        this.parentFormula = parentFormula;
        parseFormulaPartString( formulaPartString );
    }

    public FormulaPart(FormulaPart aFormulaPart) {
        this.parentFormula = aFormulaPart.parentFormula; 
        for(int i=0; i < aFormulaPart.parts.size(); i++ ){
            Object part = aFormulaPart.parts.get(i);
            if( part instanceof String ){
                parts.add(part.toString());
            }else if(part instanceof CellRef){
                CellRef cellRef = new CellRef( part.toString(), this );
                parts.add( cellRef );
                cellRefs.add( cellRef );
            }
        }
        this.defaultValue = aFormulaPart.defaultValue;
    }


    public Integer getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(Integer defaultValue) {
        this.defaultValue = defaultValue;
    }

    public void parseFormulaPartString(String formula){
        parts.clear();
        cellRefs.clear();
        formula = extractDefaultValue(formula);
        Matcher refCellMatcher = regexCellRefPattern.matcher( formula );
        int end = 0;
        CellRef cellRef = null;
        while( refCellMatcher.find() ){
            String part = formula.substring(end, refCellMatcher.start());
            part = adjustFormulaPartForCellIndex(cellRef, part);
            parts.add(part);
            cellRef = new CellRef( refCellMatcher.group(), this );
            parts.add( cellRef );
            cellRefs.add( cellRef );
            end = refCellMatcher.end();
        }
        parts.add( adjustFormulaPartForCellIndex( cellRef, formula.substring( end ) ));
    }

    private String extractDefaultValue(String formula) {
        int i = formula.indexOf(defaultValueToken);
        String resultFormula = formula;
        if(i>=0){
            resultFormula = formula.substring(0, i);
            try {
                defaultValue = Integer.valueOf( formula.substring(i+1));
            } catch (NumberFormatException e) {
                log.error("Can't parse default value constant for " + formulaPartString + " formula part. Integer expected after '@' symbol");
            }
        }
        return resultFormula;
    }

    private String adjustFormulaPartForCellIndex(CellRef cellRef, String formulaPart) {
        if( cellRef != null){
            int indStart = formulaPart.indexOf('(');
            int indEnd = formulaPart.indexOf(')');
            if( indStart == 0 && indEnd > 0){
                String cellIndex = formulaPart.substring( indStart + 1, indEnd );
                try {
                    cellRef.setCellIndex( Integer.valueOf( cellIndex ) );
                    formulaPart = formulaPart.substring( indEnd + 1 );
                } catch (NumberFormatException e) {
                    log.error("Can't parse cell index " + cellIndex + " for cell " + cellRef + ". Make sure you don't have any spaces for index part.", e);
                }
            }
        }
        return formulaPart;
    }

    void replaceCellRefs(CellRef cellRef, List rangeFormulaParts) {
        cellRefsToRemove.add( cellRef );
        for (int i = 0; i < rangeFormulaParts.size(); i++) {
            Object formulaPart = rangeFormulaParts.get(i);
            if( formulaPart instanceof CellRef ){
                cellRefsToAdd.add( formulaPart );
            }
        }
    }

    public void replaceCellRef(CellRef cellRef, List rangeFormulaParts) {
        for (int i = 0; i < parts.size(); i++) {
            Object formulaPart = parts.get(i);
            if( formulaPart == cellRef ){
                replaceFormulaPart( i, rangeFormulaParts );
                replaceCellRefs( cellRef, rangeFormulaParts );
                break;
            }
        }
    }

    private void replaceFormulaPart(int pos, List rangeFormulaParts) {
        parts.remove( pos );
        parts.addAll( pos, rangeFormulaParts );
    }

    public Collection getRefCells() {
        return cellRefs;
    }

    public String getActualFormula() {
        if( cellRefs.isEmpty() && defaultValue != null){
            return defaultValue.toString();
        }
        Object formulaPart;
        String actualFormula = "";
        for (Iterator iterator = parts.iterator(); iterator.hasNext();) {
            formulaPart =  iterator.next();
            actualFormula += formulaPart.toString();
        }

        return actualFormula;

        
    }

    public void removeCellRefs( Set cellRefsToBeRemoved ){
        List formulaPartIndexesToRemove = new ArrayList();
        Object prevFormulaPart = null;
        Object nextFormulaPart;
        for (int i = 0; i < parts.size(); i++) {
            Object formulaPart = parts.get(i);
            if( cellRefsToBeRemoved.contains( formulaPart ) ){
                formulaPartIndexesToRemove.add( new Integer( i ) );
                if( i > 0 ){
                    prevFormulaPart = parts.get( i - 1 );
                }
                if( i < parts.size() - 1 ){
                    nextFormulaPart = parts.get( i + 1 );
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
        Collections.sort( formulaPartIndexesToRemove );
        int shift = 0;
        for (int i = 0; i < formulaPartIndexesToRemove.size(); i++) {
            int index =  ((Integer) formulaPartIndexesToRemove.get(i)).intValue() ;
            parts.remove( index - shift );
            shift++;
        }
        cellRefs.removeAll( cellRefsToBeRemoved );
    }

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


    public String toString() {
        return formulaPartString;
    }
}
