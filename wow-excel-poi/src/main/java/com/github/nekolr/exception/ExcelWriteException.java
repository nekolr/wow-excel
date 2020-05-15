package com.github.nekolr.exception;

/**
 * 写异常
 */
public class ExcelWriteException extends RuntimeException {

    public ExcelWriteException() {
    }

    public ExcelWriteException(String message) {
        super(message);
    }

    public ExcelWriteException(Throwable cause) {
        super(cause);
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
