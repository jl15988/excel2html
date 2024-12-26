package com.jl15988.excel2html.stream;

/**
 * 内容流转
 *
 * @author Jalon
 * @since 2024/12/26 16:57
 **/
public class ValueStream {

    private Object value;

    public ValueStream(Object value) {
        this.value = value;
    }

    public <R> R to(Class<R> clazz) {
        return (R) this.value;
    }
}
