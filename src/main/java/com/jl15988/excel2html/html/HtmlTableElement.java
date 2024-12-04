package com.jl15988.excel2html.html;

public interface HtmlTableElement<T> extends IHtmlElement<T> {

    default String toHtmlString() {
        StringBuffer stringBuffer = new StringBuffer();

        String tagName = getTagName();
        String content = getContent();

        stringBuffer.append("<").append(tagName).append(">");
        if (content != null) {
            stringBuffer.append(content);
        } else {
            getChildrenElementList().forEach(iHtmlElement -> stringBuffer.append(iHtmlElement.toHtmlString()).append("\n"));
        }
        stringBuffer.append("</").append(tagName).append(">");
        return stringBuffer.toString();
    }
}
