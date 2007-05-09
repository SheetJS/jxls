package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public abstract class BaseBlockReader implements XLSBlockReader {
    int startRow;
    int endRow;

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }
}
