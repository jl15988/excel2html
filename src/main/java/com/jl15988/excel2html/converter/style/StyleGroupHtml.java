package com.jl15988.excel2html.converter.style;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/2 10:54
 **/
@Data
public class StyleGroupHtml {

    private String styleContent;

    private Map<String, List<String>> tagStyleUidMap;
}
