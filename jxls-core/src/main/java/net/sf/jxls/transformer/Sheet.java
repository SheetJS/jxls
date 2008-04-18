package net.sf.jxls.transformer;

import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.formula.ListRange;
import net.sf.jxls.parser.Cell;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Represents excel worksheet 
 * @author Leonid Vysochyn
 */
public class Sheet {

    Workbook workbook;

    /**
     * POI Excel workbook object
     */
    HSSFWorkbook hssfWorkbook;

    /**
     * POI Excel sheet representation
     */
    HSSFSheet hssfSheet;
    /**
     * This variable stores all list ranges found while processing template file
     */
    private Map listRanges = new HashMap();
    /**
     * Stores all named HSSFCell objects
     */
    private Map namedCells = new HashMap();

    Configuration configuration = new Configuration();

    public Sheet() {
    }

    public Sheet(HSSFWorkbook hssfWorkbook, HSSFSheet hssfSheet, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.hssfSheet = hssfSheet;
        this.configuration = configuration;
    }

    public Sheet(HSSFWorkbook hssfWorkbook, HSSFSheet hssfSheet) {
        this.hssfWorkbook = hssfWorkbook;
        this.hssfSheet = hssfSheet;
    }

    public String getSheetName(){
        return sheetName;
    }

    public void setSheetName(String sheetName){
        this.sheetName = sheetName; 
    }

    String sheetName;

    public void initSheetName(){
        for(int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++){
            HSSFSheet sheet = hssfWorkbook.getSheetAt( i );
            if( sheet == hssfSheet ){
                sheetName = hssfWorkbook.getSheetName( i );
                if( sheetName.indexOf(' ') >=0 ){
                    sheetName = "'" + sheetName + "'";
                }
            }
        }
    }

    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    public void setHssfWorkbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public void setHssfSheet(HSSFSheet hssfSheet) {
        this.hssfSheet = hssfSheet;
    }

    public HSSFSheet getHssfSheet() {
        return hssfSheet;
    }

    public Configuration getConfiguration() {
        return configuration;
    }

    public Map getListRanges() {
        return listRanges;
    }

    public Map getNamedCells() {
        return namedCells;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public void addNamedCell(String label, Cell cell){
        namedCells.put( label, cell );
    }

    public void addListRange(String name, ListRange range){
        listRanges.put( name, range );
    }

    public int getMaxColNum(){
        int maxColNum = 0;
        for(int i = hssfSheet.getFirstRowNum(); i <= hssfSheet.getLastRowNum(); i++){
            HSSFRow hssfRow = hssfSheet.getRow( i );
            if( hssfRow != null ){
                if( hssfRow.getLastCellNum() > maxColNum ){
                    maxColNum = hssfRow.getLastCellNum();
                }
            }
        }
        return maxColNum;
    }


}
