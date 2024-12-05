package com.jl15988.excel2html.html;

import java.util.*;
import java.util.stream.Collectors;

/**
 * html元素
 *
 * @author Jalon
 * @since 2024/11/30 19:51
 **/
public class HtmlElement implements IHtmlElement<HtmlElement> {

    private final String uid;

    private String id;

    private final String tagName;

    private final List<String> classList = new ArrayList<>();

    private final Map<String, Object> styleMap = new HashMap<String, Object>();

    private final Map<String, String> attrsMap = new HashMap<String, String>();

    private final List<IHtmlElement<?>> childrenList = new ArrayList<IHtmlElement<?>>();

    private String content;

    public HtmlElement(String tagName) {
        this.tagName = tagName;
        this.uid = UUID.randomUUID().toString();
    }

    @Override
    public String getUID() {
        return uid;
    }

    @Override
    public String getId() {
        return id;
    }

    @Override
    public HtmlElement setId(String id) {
        this.id = id;
        return this;
    }

    @Override
    public String getTagName() {
        return tagName;
    }

    @Override
    public List<String> getClassList() {
        return classList;
    }

    @Override
    public HtmlElement addClass(String className) {
        classList.add(className);
        return this;
    }

    @Override
    public HtmlElement addClasses(List<String> classList) {
        this.classList.addAll(classList);
        return this;
    }

    @Override
    public HtmlElement addClasses(String... classes) {
        this.classList.addAll(Arrays.asList(classes));
        return this;
    }

    @Override
    public HtmlElement removeClass(String className) {
        classList.remove(className);
        return this;
    }

    @Override
    public HtmlElement clearClass() {
        classList.clear();
        return this;
    }

    @Override
    public Map<String, Object> getStyleMap() {
        return styleMap;
    }

    @Override
    public HtmlElement setStyleMap(Map<String, Object> styleMap) {
        this.styleMap.putAll(styleMap);
        return this;
    }

    @Override
    public Object getStyle(String styleName) {
        return styleMap.get(styleName);
    }

    @Override
    public HtmlElement addStyle(String styleName, Object value) {
        styleMap.put(styleName, value);
        return this;
    }

    @Override
    public HtmlElement addStyle(ICssStyle<?> cssStyle) {
        Map<String, Object> map = cssStyle.getMap();
        styleMap.putAll(map);
        return this;
    }

    @Override
    public HtmlElement removeStyle(String styleName) {
        styleMap.remove(styleName);
        return this;
    }

    @Override
    public HtmlElement clearStyle() {
        styleMap.clear();
        return this;
    }

    @Override
    public Map<String, String> getAttributeMap() {
        return attrsMap;
    }

    @Override
    public String getAttribute(String name) {
        return attrsMap.get(name);
    }

    @Override
    public HtmlElement addAttribute(String name, String value) {
        attrsMap.put(name, value);
        return this;
    }

    @Override
    public HtmlElement removeAttribute(String name) {
        attrsMap.remove(name);
        return this;
    }

    @Override
    public HtmlElement clearAttributes() {
        attrsMap.clear();
        return this;
    }

    @Override
    public List<IHtmlElement<?>> getChildrenElementList() {
        return childrenList;
    }

    @Override
    public HtmlElement addChildElement(IHtmlElement<?> child) {
        childrenList.add(child);
        return this;
    }

    @Override
    public HtmlElement addChildElement(int index, IHtmlElement<?> child) {
        childrenList.add(index, child);
        return this;
    }

    @Override
    public HtmlElement removeChild(IHtmlElement<?> child) {
        childrenList.remove(child);
        return this;
    }

    @Override
    public HtmlElement removeChild(int index) {
        childrenList.remove(index);
        return this;
    }

    @Override
    public HtmlElement clearChildren() {
        childrenList.clear();
        return this;
    }

    @Override
    public HtmlElement setContent(String content) {
        this.content = content;
        return this;
    }

    @Override
    public String getContent() {
        return content;
    }

    @Override
    public String toHtmlString() {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("<").append(tagName);
        if (id != null && !id.isEmpty()) {
            stringBuilder.append(" ").append("id=\"").append(id).append("\"");
        }

        if (classList != null && !classList.isEmpty()) {
            stringBuilder.append(" class=\"").append(String.join(" ", classList)).append("\"");
        }

        if (styleMap != null && !styleMap.isEmpty()) {
            stringBuilder.append(" ").append("style=\"");
            styleMap.forEach((name, value) -> {
                stringBuilder.append(name).append(":").append(value).append(";");
            });
            stringBuilder.append("\"");
        }

        if (attrsMap != null && !attrsMap.isEmpty()) {
            attrsMap.forEach((name, value) -> {
                stringBuilder.append(" ").append(name).append("=\"").append(value).append("\"");
            });
        }
        stringBuilder.append(">");

        if (content != null) {
            stringBuilder.append(content);
        } else {
            stringBuilder.append("\n");
            if (childrenList != null && !childrenList.isEmpty()) {
                stringBuilder.append(childrenList.stream().map(IHtmlElement::toHtmlString).collect(Collectors.joining("\n")));
            }
        }
        stringBuilder.append("</").append(tagName).append(">");
        return stringBuilder.toString();
    }

    public static HtmlElementBuilder builder(String tagName) {
        return new HtmlElementBuilder(tagName);
    }

    public static class HtmlElementBuilder {

        private final HtmlElement htmlElement;

        public HtmlElementBuilder(String tagName) {
            this.htmlElement = new HtmlElement(tagName);
        }

        public HtmlElementBuilder id(String id) {
            htmlElement.setId(id);
            return this;
        }

        public HtmlElementBuilder addClass(String className) {
            htmlElement.addClass(className);
            return this;
        }

        public HtmlElementBuilder addClasses(List<String> classList) {
            htmlElement.addClasses(classList);
            return this;
        }

        public HtmlElementBuilder addClasses(String... classes) {
            htmlElement.addClasses(classes);
            return this;
        }

        public HtmlElementBuilder style(Map<String, Object> styleMap) {
            htmlElement.setStyleMap(styleMap);
            return this;
        }

        public HtmlElementBuilder addStyle(String styleName, Object value) {
            htmlElement.addStyle(styleName, value);
            return this;
        }

        public HtmlElementBuilder addStyle(ICssStyle<?> cssStyle) {
            htmlElement.addStyle(cssStyle);
            return this;
        }

        public HtmlElementBuilder addAttribute(String name, String value) {
            htmlElement.addAttribute(name, value);
            return this;
        }

        public HtmlElementBuilder addChildElement(IHtmlElement<?> child) {
            htmlElement.addChildElement(child);
            return this;
        }

        public HtmlElementBuilder addChildElement(int index, IHtmlElement<?> child) {
            htmlElement.addChildElement(index, child);
            return this;
        }

        public HtmlElementBuilder content(String content) {
            htmlElement.setContent(content);
            return this;
        }

        public HtmlElement build() {
            return htmlElement;
        }
    }
}
