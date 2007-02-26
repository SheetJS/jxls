package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 */
public class XLSDataReadException extends RuntimeException{

    public XLSDataReadException() {
    }

    public XLSDataReadException(String message) {
        super(message);
    }


    public XLSDataReadException(Throwable cause) {
        super(cause);
    }

    public XLSDataReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
