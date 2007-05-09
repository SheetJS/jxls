package net.sf.jxls.transformer;

import net.sf.jxls.transformation.ResultTransformation;
import net.sf.jxls.controller.SheetTransformationController;
import net.sf.jxls.tag.Block;

import java.util.Map;

/**
 * Defines row transformation methods
 */
public interface RowTransformer {
    Row getRow();
    ResultTransformation transform(SheetTransformationController stc, SheetTransformer sheetTransformer, Map beans);
    Block getTransformationBlock();
    void setTransformationBlock(Block block);
    ResultTransformation getTransformationResult();

}
