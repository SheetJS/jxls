package net.sf.jxls.transformer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.util.SheetHelper;
import net.sf.jxls.formula.FormulaController;
import net.sf.jxls.formula.FormulaControllerImpl;

import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;

/**
 * Represents excel workbook
 * @author Leonid Vysochyn
 */
public class Workbook {
    List sheets = new ArrayList();
    /**
     * POI Excel workbook object
     */
    HSSFWorkbook hssfWorkbook;

    FormulaController formulaController;

    Configuration configuration = new Configuration();

    public Workbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.configuration = configuration;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, List sheets) {
        this.hssfWorkbook = hssfWorkbook;
        this.sheets = sheets;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, List sheets, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.sheets = sheets;
        this.configuration = configuration;
    }

    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    public void setHssfWorkbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public void addSheet(Sheet sheet){
        sheets.add( sheet );
        sheet.setWorkbook( this );
    }

    public void initSheetNames(){
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            sheet.initSheetName();
        }
    }

    public Map getListRanges(){
        Map listRanges = new HashMap();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            listRanges.putAll( sheet.getListRanges() );
        }
        return listRanges;
    }

    public List findFormulas(){
        List formulas = new ArrayList();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            formulas.addAll( SheetHelper.findFormulas( sheet ) );
        }
        return formulas;
    }

    public Map createFormulaSheetMap(){
        Map formulas = new HashMap();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            formulas.put( sheet.getSheetName(), SheetHelper.findFormulas( sheet ) );
        }
        return formulas;
    }

    public FormulaController createFormulaController(){
        formulaController = new FormulaControllerImpl( this );
        return formulaController;
    }

    public FormulaController getFormulaController() {
        return formulaController;
    }


    public List getSheets() {
        return sheets;
    }

    public int getNumberOfSheets(){
        return sheets.size();
    }

    public Sheet getSheetAt(int sheetNo){
        return (Sheet) sheets.get( sheetNo );
    }

}
