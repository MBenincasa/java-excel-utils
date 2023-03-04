package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that the extension is invalid
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public class ExtensionNotValidException extends Exception {

    /**
     * Constructs an {@code ExtensionNotValidException} with {@code null}
     * as its error detail message.
     */
    public ExtensionNotValidException() {
        super();
    }

    /**
     * Constructs an {@code ExtensionNotValidException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public ExtensionNotValidException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExtensionNotValidException} with the specified detail message
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
    public ExtensionNotValidException(String message, Throwable cause) {
        super(message, cause);
    }
}
