package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that the Excel row was not found
 * @author Mirko Benincasa
 * @since 0.4.1
 */
public class RowNotFoundException extends Exception {

    /**
     * Constructs an {@code RowNotFoundException} with {@code null}
     * as its error detail message.
     */
    public RowNotFoundException() {
        super();
    }

    /**
     * Constructs an {@code RowNotFoundException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public RowNotFoundException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code RowNotFoundException} with the specified detail message
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
    public RowNotFoundException(String message, Throwable cause) {
        super(message, cause);
    }
}
