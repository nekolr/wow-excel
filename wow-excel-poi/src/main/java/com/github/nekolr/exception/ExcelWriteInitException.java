package com.github.nekolr.exception;

/**
 * 写初始化异常
 */
public class ExcelWriteInitException extends ExcelWriteException {

    public ExcelWriteInitException() {
    }

    public ExcelWriteInitException(String message) {
        super(message);
    }

    public ExcelWriteInitException(Throwable cause) {
        super(cause);
    }

    public ExcelWriteInitException(String message, Throwable cause) {
        super(message, cause);
    }
}
