package com.jl15988.excel2html.wander;

/**
 * 流转抽象
 *
 * @author Jalon
 * @since 2024/12/27 9:01
 **/
public abstract class WanderBase {

    protected Object value;

    protected Throwable throwable;

    protected boolean done;

    public void set(Object value) {
        this.value = value;
    }

    public void setThrowable(Throwable throwable) {
        this.throwable = throwable;
    }

    public void setDone() {
        this.done = true;
    }
}
