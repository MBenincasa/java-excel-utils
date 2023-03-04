package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that there was an error reading a cell value
 * @author Mirko Benincasa
 * @since 0.3.0
 */
public class ReadValueException extends Exception {

    /**
     * Constructs an {@code ReadValueException} with {@code null}
     * as its error detail message.
     */
    public ReadValueException() {
        super();
    }

    /**
     * Constructs an {@code ReadValueException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public ReadValueException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ReadValueException} with the specified detail message
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
    public ReadValueException(String message, Throwable cause) {
        super(message, cause);
    }
}
