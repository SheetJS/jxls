package net.sf.jxls.controller;

/**
 * This class is represents a single cell in a transformation matrix
 * @author Leonid Vysochyn
 */
public class MatrixCell {
    int rowNum;
    int colNum;

    public MatrixCell() {
    }

    public MatrixCell(int rowNum, int colNum) {
        this.rowNum = rowNum;
        this.colNum = colNum;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getColNum() {
        return colNum;
    }

    public void setColNum(int colNum) {
        this.colNum = colNum;
    }

    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        final MatrixCell that = (MatrixCell) o;

        if (colNum != that.colNum) return false;
        if (rowNum != that.rowNum) return false;

        return true;
    }

    public int hashCode() {
        int result;
        result = rowNum;
        result = 29 * result + colNum;
        return result;
    }

    public String toString() {
        return "(" + rowNum + ", " + colNum + ")";
    }
}
