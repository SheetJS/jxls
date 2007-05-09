package net.sf.jxls.exception;

/**
 * @author Leonid Vysochyn
 */
public class TaglibRegistrationException extends RuntimeException{
    public TaglibRegistrationException(String message) {
        super(message);
    }

    public TaglibRegistrationException(String message, Throwable cause) {
        super(message, cause);
    }

    public TaglibRegistrationException(Throwable cause) {
        super(cause);
    }

    public TaglibRegistrationException() {
    }
}
