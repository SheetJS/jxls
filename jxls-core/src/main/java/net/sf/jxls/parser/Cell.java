package net.sf.jxls.parser;

import java.util.ArrayList;
import java.util.List;

import net.sf.jxls.formula.Formula;
import net.sf.jxls.tag.Tag;
import net.sf.jxls.transformer.Row;
import net.sf.jxls.transformer.RowCollection;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.Region;

/**
 * Represents excel cell
 * @author Leonid Vysochyn
 */
public class Cell {

    private Row row;
    private Property collectionProperty;
    private HSSFCell hssfCell;

    private Formula formula;
    private String label;
    private int dependentRowNumber;

    private String collectionName;

    private String hssfCellValue;
    private String stringCellValue;
    private String metaInfo;

    private Region mergedRegion;

    private List expressions = new ArrayList();

    private Tag tag;

    private RowCollection rowCollection;


    public Cell(HSSFCell hssfCell, Row row) {
        this.setHssfCell(hssfCell);
        this.setRow(row);
    }


    public Tag getTag() {
        return tag;
    }

    public void setTag(Tag tag) {
        this.tag = tag;
    }

    public List getExpressions() {
        return expressions;
    }

    public void setExpressions(List expressions) {
        this.expressions = expressions;
    }

    public Region getMergedRegion() {
        return mergedRegion;
    }

    public int getDependentRowNumber() {
        return dependentRowNumber;
    }


    public RowCollection getRowCollection() {
        return rowCollection;
    }

    public void setRowCollection(RowCollection rowCollection) {
        this.rowCollection = rowCollection;
    }

    public String getCollectionName() {
        return collectionName;
    }

    public Formula getFormula() {
        return formula;
    }

    public void setFormula(Formula formula) {
        this.formula = formula;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public Property getCollectionProperty() {
        return collectionProperty;
    }

    public void setCollectionProperty(Property collectionProperty) {
        this.collectionProperty = collectionProperty;
    }

    public HSSFCell getHssfCell() {
        return hssfCell;
    }

    public void setHssfCell(HSSFCell hssfCell) {
        this.hssfCell = hssfCell;
    }

    public String toCellName() {
        CellReference cellRef = new CellReference(getRow().getHssfRow().getRowNum(), getHssfCell().getCellNum());
        return cellRef.toString();
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public boolean isFormula(){
        return getFormula() !=null;
    }


    public String getHssfCellValue() {
        return hssfCellValue;
    }

    public String getStringCellValue() {
        return stringCellValue;
    }

    public String getMetaInfo() {
        return metaInfo;
    }

    public boolean isEmpty(){
        return getHssfCellValue() == null || getHssfCellValue().length() == 0 || getHssfCell().getCellType() == HSSFCell.CELL_TYPE_BLANK;
    }

    public boolean isNull(){
        return getHssfCell() == null;
    }

    public String toString() {
        return getHssfCellValue();
    }

    public void setDependentRowNumber(int dependentRowNumber) {
        this.dependentRowNumber = dependentRowNumber;
    }

    public void setCollectionName(String collectionName) {
        this.collectionName = collectionName;
    }

    public void setHssfCellValue(String hssfCellValue) {
        this.hssfCellValue = hssfCellValue;
    }

    public void setStringCellValue(String stringCellValue) {
        this.stringCellValue = stringCellValue;
    }

    public void setMetaInfo(String metaInfo) {
        this.metaInfo = metaInfo;
    }

    public void setMergedRegion(Region mergedRegion) {
        this.mergedRegion = mergedRegion;
    }
}
