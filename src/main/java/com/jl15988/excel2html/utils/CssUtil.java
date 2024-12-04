package com.jl15988.excel2html.utils;

import java.util.Map;
import java.util.regex.Pattern;

/**
 * @author Jalon
 * @since 2024/12/2 16:21
 **/
public class CssUtil {

    public static String randomName(int length) {
        return RandomUtil.randomHash2(length);
    }

    public static String randomName(int length, String suffix) {
        return RandomUtil.randomHash2(length) + "_" + suffix;
    }

    /**
     * 压缩样式
     *
     * @param css 样式
     */
    public static String compress(String css) {
//        // 移除注释
        css = Pattern.compile("/\\*[^*]*\\*+(?:[^/*][^*]*\\*+)*/", Pattern.CASE_INSENSITIVE | Pattern.MULTILINE | Pattern.DOTALL).matcher(css).replaceAll("");
        // 移除多余的空白符
        css = Pattern.compile("\\s+", Pattern.CASE_INSENSITIVE | Pattern.MULTILINE | Pattern.DOTALL).matcher(css).replaceAll(" ");
        // 移除选择器与花括号之间和花括号内的多余空白
        css = css.replaceAll("\\s+(?![^}]*\\})", "");
        css = css.replaceAll("\\s+(:\\s+)(?![^;]*;)", "$1");
        return css.trim();
    }

    /**
     * Map 转 css 字符串
     *
     * @param map Map
     */
    public static String mapToCssString(Map<String, Object> map) {
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            String key = entry.getKey();
            String value = String.valueOf(entry.getValue());
            sb.append(key).append(":").append(value).append("; ");
        }
        return sb.toString();
    }
}
