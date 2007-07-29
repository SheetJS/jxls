package net.sf.jxls.reader;

/**
 * @author Leonid Vysochyn
 * @version 1.0 Jul 30, 2007
 */
public class XLSReadMessage {
    String message;
    Exception exception;

    public XLSReadMessage(String message, Exception exception) {
        this.message = message;
        this.exception = exception;
    }


    public XLSReadMessage(String message) {
        this.message = message;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public Exception getException() {
        return exception;
    }

    public void setException(Exception exception) {
        this.exception = exception;
    }
}
