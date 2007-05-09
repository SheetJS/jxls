package net.sf.jxls.controller;

import net.sf.jxls.controller.WorkbookTransformationController;
import net.sf.jxls.transformer.Workbook;

import java.util.List;
import java.util.ArrayList;

import net.sf.jxls.controller.SheetTransformationController;

/**
 * Simple implementation of {@link WorkbookTransformationController} based on the list of SheetTransformationControllers
 * @author Leonid Vysochyn
 */
public class WorkbookTransformationControllerImpl implements WorkbookTransformationController {
    List sheetTransformationControllers = new ArrayList();

    Workbook workbook;

    public WorkbookTransformationControllerImpl(Workbook hssfWorkbook) {
        this.workbook = hssfWorkbook;
    }

    public List getSheetTransformationControllers() {
        return sheetTransformationControllers;
    }

    public void setSheetTransformationControllers(List sheetTransformationControllers) {
        this.sheetTransformationControllers = sheetTransformationControllers;
    }

    public void addSheetTransformationController(SheetTransformationController sheetTransformationController) {
        sheetTransformationControllers.add( sheetTransformationController );
    }

    public void removeSheetTransformationController(SheetTransformationController sheetTransformationController) {
        sheetTransformationControllers.remove( sheetTransformationController );
    }

    public Workbook getWorkbook() {
        return workbook;
    }

}
