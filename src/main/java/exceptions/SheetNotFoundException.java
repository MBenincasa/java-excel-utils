package exceptions;

public class SheetNotFoundException extends Exception {

    public SheetNotFoundException() {
        super();
    }

    public SheetNotFoundException(String message) {
        super(message);
    }
}
