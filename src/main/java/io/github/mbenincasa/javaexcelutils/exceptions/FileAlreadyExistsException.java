package io.github.mbenincasa.javaexcelutils.exceptions;

/**
 * This exception signals that the file already exists
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public class FileAlreadyExistsException extends Exception {

    /**
     * Constructs an {@code FileAlreadyExistsException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public FileAlreadyExistsException(String message) {
        super(message);
    }
}
