package com.easyexcel.test.common.excel.exception;

/**
 * Created Date by 2020/5/9 0009.
 *
 * Excel 异常类
 * @author Parker
 */
public class ExcelException extends Exception{

    /**
     * Constructs a new exception with the specified detail message.  The
     * cause is not initialized, and may subsequently be initialized by
     * a call to {@link #initCause}.
     *
     * @param message the detail message. The detail message is saved for
     *                later retrieval by the {@link #getMessage()} method.
     */
    public ExcelException(String message) {
        super(message);
    }

}
