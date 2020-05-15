package com.github.nekolr.exception;

/**
 * 初始化异常，比如实体创建失败、注解配置不规范等
 */
public class ExcelReadInitException extends ExcelReadException {

    public ExcelReadInitException() {
    }

    public ExcelReadInitException(String message) {
        super(message);
    }

    public ExcelReadInitException(Throwable cause) {
        super(cause);
    }

    public ExcelReadInitException(String message, Throwable cause) {
        super(message, cause);
    }
}
