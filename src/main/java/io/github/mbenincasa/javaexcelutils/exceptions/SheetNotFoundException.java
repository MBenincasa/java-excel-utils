package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that the Excel sheet was not found
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public class SheetNotFoundException extends Exception {

    /**
     * Constructs an {@code SheetNotFoundException} with {@code null}
     * as its error detail message.
     */
    public SheetNotFoundException() {
        super();
    }

    /**
     * Constructs an {@code SheetNotFoundException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public SheetNotFoundException(String message) {
        super(message);
    }
}
