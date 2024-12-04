package com.jl15988.excel2html.html;

import java.util.*;

/**
 * html元素构建器
 *
 * @author Jalon
 * @since 2024/11/29 17:32
 **/
public class HtmlElementBase2 {

    private String name;

    private String content;

    private List<String> classList;

    private String id;

    private Map<String, Object> style;

    private Map<String, String> attrs = new HashMap<String, String>();

    private List<HtmlElementBase2> children = new ArrayList<HtmlElementBase2>();

    public HtmlElementBase2(String name) {
        this.name = name;
    }

    /**
     * 是否为文本
     */
    public boolean isText() {
        return "text".equals(name);
    }

    /**
     * 设置内容
     *
     * @param content 会覆盖所有子元素
     */
    public HtmlElementBase2 setContent(String content) {
        this.content = content;
        return this;
    }

    /**
     * 获取内容
     */
    public String getContent() {
        return content;
    }

    /**
     * 添加属性
     *
     * @param name  属性名
     * @param value 属性值
     */
    public HtmlElementBase2 attr(String name, String value) {
        this.attrs.put(name, value);
        return this;
    }

    /**
     * 添加子元素
     *
     * @param child 子元素
     */
    public HtmlElementBase2 add(HtmlElementBase2 child) {
        this.children.add(child);
        return this;
    }

    /**
     * 添加文本
     *
     * @param text 文本
     */
    public HtmlElementBase2 add(String text) {
        HtmlElementBase2 textEl = new HtmlElementBase2("text");
        textEl.setContent(text);
        this.children.add(textEl);
        return this;
    }

    private String buildAttr() {
        StringBuilder attrString = new StringBuilder();
        this.attrs.forEach((k, v) -> {
            attrString.append(" ").append(k).append("=\"").append(v).append("\"");
        });
        return attrString.toString();
    }

    private String buildChildren() {
        StringBuilder childrenString = new StringBuilder();
        this.children.forEach(child -> {
            String content = child.getContent();
            childrenString.append("\n").append(Optional.of(content).orElse(child.toHtmlString()));
        });
        return childrenString.toString();
    }

    public String toHtmlString() {
        StringBuilder elString = new StringBuilder();
        elString.append("<")
                .append(name)
                .append(">")
                .append(buildAttr())
                .append(buildChildren())
                .append("</")
                .append(name)
                .append(">");
        return elString.toString();
    }
}
