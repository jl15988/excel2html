package com.jl15988.excel2html.html;

import com.jl15988.excel2html.utils.CssUtil;

import java.util.Map;
import java.util.Objects;

/**
 * @author Jalon
 * @since 2024/11/29 17:06
 **/
public interface ICssStyle<T> {

    T set(String name, Object value);

    default T set(boolean condition, String name, Object value) {
        if (condition) {
            return set(name, value);
        }
        return (T) this;
    }

    default T setIfExists(String name, Object value) {
        if (Objects.nonNull(value)) {
            return set(name, value);
        }
        return (T) this;
    }

    Object get(String name);

    Object getOrDefault(String name, Object defaultValue);

    Map<String, Object> getMap();

    default String toHtmlString() {
        Map<String, Object> styleMap = this.getMap();
        return CssUtil.mapToCssString(styleMap);
    }
}
