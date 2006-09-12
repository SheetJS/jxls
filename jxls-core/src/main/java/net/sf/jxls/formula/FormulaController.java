package net.sf.jxls.formula;

import net.sf.jxls.transformation.BlockTransformation;

import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public interface FormulaController {
    public void updateWorkbookFormulas(BlockTransformation transformation);
    public Map getSheetFormulasMap();

    void writeFormulas(FormulaResolver formulaResolver);
}
