package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that the file header is missing
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public class HeaderNotPresentException extends Exception {

    /**
     * Constructs an {@code HeaderNotPresentException} with {@code null}
     * as its error detail message.
     */
    public HeaderNotPresentException() {
        super();
    }

    /**
     * Constructs an {@code HeaderNotPresentException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public HeaderNotPresentException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code HeaderNotPresentException} with the specified detail message
     * and cause.
     *
     * <p> Note that the detail message associated with {@code cause} is
     * <i>not</i> automatically incorporated into this exception's detail
     * message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     *
     * @param cause
     *        The cause (which is saved for later retrieval by the
     *        {@link #getCause()} method).  (A null value is permitted,
     *        and indicates that the cause is nonexistent or unknown.)
     */
    public HeaderNotPresentException(String message, Throwable cause) {
        super(message, cause);
    }
}
