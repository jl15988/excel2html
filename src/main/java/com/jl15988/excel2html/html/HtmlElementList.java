package com.jl15988.excel2html.html;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Jalon
 * @since 2024/11/29 17:52
 **/
public class HtmlElementList {

    private final List<HtmlElement> elementList = new ArrayList<>();

    public HtmlElementList add(HtmlElement element) {
        this.elementList.add(element);
        return this;
    }

    public HtmlElementList remove(HtmlElement element) {
        this.elementList.remove(element);
        return this;
    }

    public String toHtmlString() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < elementList.size(); i++) {
            sb.append(elementList.get(i).toHtmlString());
            if (i == elementList.size() - 1) {
                sb.append("\n");
            }
        }
        return sb.toString();
    }
}
