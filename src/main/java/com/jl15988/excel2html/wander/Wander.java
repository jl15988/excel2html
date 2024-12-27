package com.jl15988.excel2html.wander;

import java.lang.reflect.Method;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 流转
 *
 * @author Jalon
 * @since 2024/12/27 8:56
 **/
public interface Wander {

    Object get();

    boolean isDone();

    void setDone();

    void exeJudge(Object... args);

    boolean testCondition(Function<Method, Boolean> conditionMapper);

    default <R> Wander judge(boolean condition, Object... args) {
        if (!this.isDone() && condition) {
            this.exeJudge(args);
            this.setDone();
        }
        return this;
    }

    default <R> Wander judge(Function<Method, Boolean> conditionMapper, Object... args) {
        if (!this.isDone() && this.testCondition(conditionMapper)) {
            this.exeJudge(args);
            this.setDone();
        }
        return this;
    }

    Wander errorHandle(Consumer<Throwable> throwableConsumer);

    default <R> R to(Class<R> clazz) {
        Object value = this.get();
        if (value == null) return null;
        return (R) value;
    }
}
