package exceptions;

public class HeaderNotPresentException extends Exception {

    public HeaderNotPresentException() {
        super();
    }

    public HeaderNotPresentException(String message) {
        super(message);
    }

    public HeaderNotPresentException(String message, Throwable cause) {
        super(message, cause);
    }
}
