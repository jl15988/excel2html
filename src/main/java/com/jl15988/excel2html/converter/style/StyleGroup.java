package com.jl15988.excel2html.converter.style;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/2 10:32
 **/
@Data
public class StyleGroup {

    private String tagUid;

    private String styleUid;

    private Map<String, Object> styleMap = new HashMap<String, Object>();

    public void addStyle(String name, String value) {
        styleMap.put(name, value);
    }
}
