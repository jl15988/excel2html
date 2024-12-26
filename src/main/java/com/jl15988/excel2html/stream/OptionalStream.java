package com.jl15988.excel2html.stream;

import java.util.function.Function;

/**
 * 可选流转
 *
 * @author Jalon
 * @since 2024/12/26 17:01
 **/
public class OptionalStream<T> {

    private T obj;

    private Object value;

    public OptionalStream(T obj) {
        this.obj = obj;
    }

    public <R> OptionalStream<T> judge(boolean condition, Function<T, R> mapper) {
        if (condition) {
            R res = mapper.apply(this.obj);
            this.value = res;
        }
        return this;
    }

    public <R> OptionalStream<T> judge(Function<T, Boolean> conditionMapper, Function<T, R> mapper) {
        if (conditionMapper.apply(this.obj)) {
            R res = mapper.apply(this.obj);
            this.value = res;
        }
        return this;
    }

    public <R> R to(Class<R> clazz) {
        if (this.value == null) return null;
        return (R) this.value;
    }

    public Object get() {
        return this.value;
    }
}
