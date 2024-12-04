package com.jl15988.excel2html.model.style;

import com.jl15988.excel2html.html.ICssStyle;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/11/29 17:07
 **/
public class FontICssStyle implements ICssStyle<FontICssStyle> {

    private final Map<String, Object> styleMap = new HashMap<String, Object>();

    @Override
    public FontICssStyle set(String name, Object value) {
        styleMap.put(name, value);
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
