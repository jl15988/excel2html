package com.jl15988.excel2html.html;

import java.util.HashMap;
import java.util.Map;

/**
 * html meta
 *
 * @author Jalon
 * @since 2024/12/1 12:12
 **/
public class HtmlLink {

    private String rel;

    private String href;

    private final Map<String, String> attrs = new HashMap<String, String>();

    public String getRel() {
        return rel;
    }

    public HtmlLink setRel(String rel) {
        this.rel = rel;
        return this;
    }

    public String getHref() {
        return href;
    }

    public HtmlLink setHref(String href) {
        this.href = href;
        return this;
    }

    public Map<String, String> getAttrs() {
        return attrs;
    }

    public String getAttr(String name) {
        return attrs.get(name);
    }

    public HtmlLink setAttr(String name, String value) {
        attrs.put(name, value);
        return this;
    }

    public HtmlLink removeAttr(String name) {
        attrs.remove(name);
        return this;
    }

    public HtmlLink clearAttrs() {
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
        stringBuilder.append("<link");
        attrs.forEach((k, v) -> {
            stringBuilder.append(k).append("=").append("\"").append(v).append("\"").append(" ");
        });
        stringBuilder.append(">");
        return stringBuilder.toString();
    }
}
