package exceptions;

public class OpenWorkbookException extends Exception {

    public OpenWorkbookException() {
        super();
    }

    public OpenWorkbookException(String message) {
        super(message);
    }

    public OpenWorkbookException(String message, Throwable cause) {
        super(message, cause);
    }
}
