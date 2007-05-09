package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public class XLSDataReadException extends RuntimeException{

    String cellName;

    public XLSDataReadException() {
    }

    public XLSDataReadException(String message) {
        super(message);
    }


    public XLSDataReadException(String cellName, String message, Throwable cause) {
        super(message, cause);
        this.cellName = cellName;
    }

    public XLSDataReadException(Throwable cause) {
        super(cause);
    }

    public XLSDataReadException(String message, Throwable cause) {
        super(message, cause);
    }


    public String getCellName() {
        return cellName;
    }

    public void setCellName(String cellName) {
        this.cellName = cellName;
    }
}
