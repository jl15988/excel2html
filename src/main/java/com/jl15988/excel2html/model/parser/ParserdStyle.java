package com.jl15988.excel2html.model.parser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/5 10:53
 **/
public class ParserdStyle {

    public Map<String, Object> styleMap;

    public List<String> classList;

    public ParserdStyle() {
        styleMap = new HashMap<String, Object>();
        classList = new ArrayList<String>();
    }

    public void merge(ParserdStyle style) {
        styleMap.putAll(style.styleMap);
        classList.addAll(style.classList);
    }
}
