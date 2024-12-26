package com.jl15988.excel2html.stream;

import java.util.Objects;

/**
 * 流转条件
 *
 * @author Jalon
 * @since 2024/12/26 17:22
 **/
public class StreamCondition {

    public Object value;

    public StreamCondition(Object value) {
        this.value = value;
    }

    public boolean equal(Object target) {
        if (Objects.isNull(this.value)) {
            return Objects.isNull(target);
        }
        return this.value.equals(target);
    }
}
