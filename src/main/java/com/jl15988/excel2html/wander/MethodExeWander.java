package com.jl15988.excel2html.wander;

import java.lang.reflect.Method;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 方法执行流转
 *
 * @author Jalon
 * @since 2024/12/27 9:01
 **/
public class MethodExeWander extends WanderBase implements Wander {

    private final Class<?> clazz;

    private final Method method;

    public MethodExeWander(Class<?> clazz, Method method) {
        this.clazz = clazz;
        this.method = method;
    }

    @Override
    public Object get() {
        return this.value;
    }

    @Override
    public boolean isDone() {
        return this.done;
    }

    @Override
    public void exeJudge(Object... args) {
        if (Objects.isNull(this.method)) return;
        try {
            this.value = method.invoke(clazz, args);
        } catch (Throwable e) {
            this.throwable = e;
        }
    }

    @Override
    public boolean testCondition(Function<Method, Boolean> conditionMapper) {
        return conditionMapper.apply(this.method);
    }

    @Override
    public MethodExeWander errorHandle(Consumer<Throwable> throwableConsumer) {
        throwableConsumer.accept(this.throwable);
        return this;
    }
}
