package com.jl15988.excel2html.enums;

import java.util.Locale;

/**
 * 通用元素 class
 *
 * @author Jalon
 * @since 2024/12/24 15:54
 **/
public enum CommonElementClass {

    VALUE_END_SPACES;

    public String value() {
        return this.name().replaceAll("_", "-").toLowerCase(Locale.ROOT);
    }
}
