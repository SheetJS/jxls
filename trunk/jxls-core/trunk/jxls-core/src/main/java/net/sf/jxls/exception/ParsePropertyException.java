package net.sf.jxls.exception;

/**
 * This exception is thrown by {@link net.sf.jxls.transformer.XLSTransformer} when problems with dynamic properties access occur
 *
 * @author Leonid Vysochyn
 */
public class ParsePropertyException extends RuntimeException {
    public ParsePropertyException() {
    }

    public ParsePropertyException(String message) {
        super(message);
    }
}
