package exceptions;

public class ExtensionNotValidException extends Exception {

    public ExtensionNotValidException() {
        super();
    }

    public ExtensionNotValidException(String message) {
        super(message);
    }

    public ExtensionNotValidException(String message, Throwable cause) {
        super(message, cause);
    }
}
