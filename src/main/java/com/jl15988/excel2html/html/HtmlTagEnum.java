package com.jl15988.excel2html.html;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author Jalon
 * @since 2024/12/1 19:55
 **/
@Getter
@AllArgsConstructor
public enum HtmlTagEnum {

    HTML,
    HEAD,
    TITLE,
    STYLE,
    BODY,
    LINK,
    META,
    TABLE,
    THEAD,
    TR,
    TD,
    DIV;

    public String getName() {
        return name().toLowerCase();
    }
}
