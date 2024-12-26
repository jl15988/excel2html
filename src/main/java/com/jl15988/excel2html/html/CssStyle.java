package com.jl15988.excel2html.html;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/3 16:02
 **/
public class CssStyle implements ICssStyle<CssStyle> {

    private final Map<String, Object> styleMap = new HashMap<String, Object>();

    @Override
    public CssStyle set(String name, Object value) {
        styleMap.put(name, value);
        return this;
    }

    @Override
    public CssStyle addAll(Map<String, Object> styleMap) {
        this.styleMap.putAll(styleMap);
        return this;
    }

    @Override
    public Object get(String name) {
        return styleMap.get(name);
    }

    @Override
    public Object getOrDefault(String name, Object defaultValue) {
        return styleMap.getOrDefault(name, defaultValue);
    }

    @Override
    public Map<String, Object> getMap() {
        return this.styleMap;
    }
}
