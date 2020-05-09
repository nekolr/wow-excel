package com.nekolr.exception;

/**
 * Excel 读过程产生的异常
 */
public class ExcelReadException extends RuntimeException {

    public ExcelReadException() {

    }

    public ExcelReadException(String message) {
        super(message);
    }

    public ExcelReadException(Throwable cause) {
        super(cause);
    }

    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
