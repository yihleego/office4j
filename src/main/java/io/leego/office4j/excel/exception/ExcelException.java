package io.leego.office4j.excel.exception;

/**
 * @author Yihleego
 */
public class ExcelException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    /**
     * Constructs an <code>ExcelException</code> with no detail message.
     */
    public ExcelException() {
        super();
    }

    /**
     * Constructs an <code>ExcelException</code> with the specified detail message.
     * @param message detail message
     */
    public ExcelException(String message) {
        super(message);
    }

    /**
     * Constructs an <code>ExcelException</code> with the specified detail message and cause.
     * @param message detail message
     * @param cause   the cause
     */
    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an <code>ExcelException</code> with the cause.
     * @param cause the cause
     */
    public ExcelException(Throwable cause) {
        super(cause);
    }

}