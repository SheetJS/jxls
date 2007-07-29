package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public abstract class BaseBlockReader implements XLSBlockReader {
    int startRow;
    int endRow;

    XLSReadStatus readStatus = new XLSReadStatus();

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

    public XLSReadStatus getReadStatus() {
        return readStatus;
    }

    public void setReadStatus(XLSReadStatus readStatus) {
        this.readStatus = readStatus;
    }
}
