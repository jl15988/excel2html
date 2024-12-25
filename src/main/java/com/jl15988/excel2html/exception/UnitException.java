package com.jl15988.excel2html.exception;

/**
 * 单位异常
 *
 * @author Jalon
 * @since 2024/12/25 14:17
 **/
public class UnitException extends RuntimeException {

    public UnitException(String message, Throwable cause) {
        super(message, cause);
    }

    public UnitException(String message) {
        super(message);
    }
}
