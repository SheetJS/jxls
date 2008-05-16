package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public class XLSDataReadException extends RuntimeException{
	private static final long serialVersionUID = 1L;

	String cellName;

    XLSReadStatus readStatus;

    public XLSDataReadException() {
    }

    public XLSDataReadException(String message) {
        super(message);
    }


    public XLSDataReadException(String cellName, String message, Throwable cause) {
        super(message, cause);
        this.cellName = cellName;
    }

    public XLSDataReadException(String cellName, String message, XLSReadStatus status) {
        super(message);
        this.cellName = cellName;
        this.readStatus = status;
    }



    public XLSDataReadException(Throwable cause) {
        super(cause);
    }

    public XLSDataReadException(String message, Throwable cause) {
        super(message, cause);
    }


    public XLSDataReadException(String message, XLSReadStatus readStatus) {
        super(message);
        this.readStatus = readStatus;
    }

    public String getCellName() {
        return cellName;
    }

    public void setCellName(String cellName) {
        this.cellName = cellName;
    }


    public XLSReadStatus getReadStatus() {
        return readStatus;
    }

    public void setReadStatus(XLSReadStatus readStatus) {
        this.readStatus = readStatus;
    }
}
