package com.jl15988.excel2html.html;

import java.util.HashMap;
import java.util.Map;

/**
 * html meta
 *
 * @author Jalon
 * @since 2024/12/1 12:12
 **/
public class HtmlMeta {

    private final Map<String, String> attrs = new HashMap<String, String>();

    public Map<String, String> getAttrs() {
        return attrs;
    }

    public String getAttr(String name) {
        return attrs.get(name);
    }

    public HtmlMeta setAttr(String name, String value) {
        attrs.put(name, value);
        return this;
    }

    public HtmlMeta removeAttr(String name) {
        attrs.remove(name);
        return this;
    }

    public HtmlMeta clearAttrs() {
        attrs.clear();
        return this;
    }

    public boolean hasAttr(String name) {
        return attrs.containsKey(name);
    }

    public boolean hasAttr(String name, String value) {
        if (!attrs.containsKey(name)) {
            return false;
        }
        return attrs.get(name).equals(value);
    }

    public String toHtmlString() {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("<meta ");
        attrs.forEach((k, v) -> {
            stringBuilder.append(k).append("=").append("\"").append(v).append("\"").append(" ");
        });
        stringBuilder.append(">");
        return stringBuilder.toString();
    }

    public static HtmlMetaBuilder builder() {
        return new HtmlMetaBuilder();
    }

    public static class HtmlMetaBuilder {

        private final HtmlMeta htmlMeta;

        public HtmlMetaBuilder() {
            htmlMeta = new HtmlMeta();
        }

        public HtmlMetaBuilder addAttr(String name, String value) {
            htmlMeta.attrs.put(name, value);
            return this;
        }

        public HtmlMeta build() {
            return htmlMeta;
        }
    }
}
