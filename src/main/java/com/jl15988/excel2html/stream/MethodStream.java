package com.jl15988.excel2html.stream;

import java.lang.reflect.Method;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 方法流转
 *
 * @author Jalon
 * @since 2024/12/26 17:42
 **/
public class MethodStream {

    private final Class<?> clazz;

    private final Method method;

    private Object value;

    private Throwable throwable;

    public MethodStream(Class<?> clazz, Method method) {
        this.clazz = clazz;
        this.method = method;
    }

    private void exeJudge(Object... args) {
        try {
            this.value = method.invoke(clazz, args);
        } catch (Throwable e) {
            this.throwable = e;
        }
    }

    public <R> MethodStream judge(boolean condition, Object... args) {
        if (condition) {
            this.exeJudge(args);
        }
        return this;
    }

    public <R> MethodStream judge(Function<Method, Boolean> conditionMapper, Object... args) {
        if (conditionMapper.apply(this.method)) {
            this.exeJudge(args);
        }
        return this;
    }

    public MethodStream errorHandle(Consumer<Throwable> throwableConsumer) {
        throwableConsumer.accept(this.throwable);
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
