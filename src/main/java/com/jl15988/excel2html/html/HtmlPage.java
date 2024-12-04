package com.jl15988.excel2html.html;

import com.jl15988.excel2html.utils.CssUtil;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Jalon
 * @since 2024/12/1 12:10
 **/
public class HtmlPage {

    private String lang;

    private String title;

    private final List<HtmlMeta> metaList = new ArrayList<>();

    private final List<HtmlLink> linkList = new ArrayList<>();

    private final List<HtmlElement> elementList = new ArrayList<>();

    private final List<HtmlStyle> styleList = new ArrayList<>();

    private List<String> styleContentList = new ArrayList<>();

    private boolean isHasHtmlContainer = true;

    private boolean isHasHtmlStyle = true;

    public String getLang() {
        return lang;
    }

    public HtmlPage setLang(String lang) {
        this.lang = lang;
        return this;
    }

    public String getTitle() {
        return title;
    }

    public HtmlPage setTitle(String title) {
        this.title = title;
        return this;
    }

    public List<HtmlMeta> getMetaList() {
        return metaList;
    }

    public HtmlPage addMeta(HtmlMeta meta) {
        metaList.add(meta);
        return this;
    }

    public List<HtmlLink> getLinkList() {
        return linkList;
    }

    public HtmlPage addLink(HtmlLink link) {
        linkList.add(link);
        return this;
    }

    public List<HtmlElement> getElementList() {
        return elementList;
    }

    public HtmlPage addElement(HtmlElement element) {
        elementList.add(element);
        return this;
    }

    public HtmlPage addElements(List<HtmlElement> elements) {
        elementList.addAll(elements);
        return this;
    }

    public HtmlPage removeElement(HtmlElement element) {
        elementList.remove(element);
        return this;
    }

    public HtmlPage clearElements() {
        elementList.clear();
        return this;
    }

    public List<HtmlStyle> getStyleList() {
        return styleList;
    }

    public HtmlPage addStyle(HtmlStyle style) {
        styleList.add(style);
        return this;
    }

    public HtmlPage removeStyle(HtmlStyle style) {
        styleList.remove(style);
        return this;
    }

    public HtmlPage clearStyles() {
        styleList.clear();
        return this;
    }

    public HtmlPage addStyleContent(String styleContent) {
        styleContentList.add(styleContent);
        return this;
    }

    public HtmlPage removeStyleContent(String styleContent) {
        styleContentList.remove(styleContent);
        return this;
    }

    public HtmlPage clearStyleContents() {
        styleContentList.clear();
        return this;
    }

    public HtmlPage setHasHtmlContainer(boolean hasHtmlContainer) {
        this.isHasHtmlContainer = hasHtmlContainer;
        return this;
    }

    public HtmlPage setHasHtmlStyle(boolean hasHtmlStyle) {
        this.isHasHtmlStyle = hasHtmlStyle;
        return this;
    }

    public String getStyleHtmlString() {
        StringBuilder stringBuilder = new StringBuilder();
        if (!styleContentList.isEmpty()) {
            styleContentList.forEach(style -> stringBuilder.append("<style type=\"text/css\" rel=\"stylesheet\">").append(CssUtil.compress(style)).append("</style>").append("\n"));
        } else {
            styleList.forEach(style -> stringBuilder.append(style.toHtmlString()).append("\n"));
        }
        return stringBuilder.toString();
    }

    public String toHtmlString() {
        StringBuilder stringBuilder = new StringBuilder();
        if (this.isHasHtmlContainer) {
            stringBuilder.append("<!DOCTYPE html>\n");
            stringBuilder.append("<html");
            if (lang != null) {
                stringBuilder.append(" lang=\"").append(lang).append("\"");
            }
            stringBuilder.append(">\n");
            stringBuilder.append("<head>\n");

            stringBuilder.append("<title>").append(title == null ? "" : title).append("</title>\n");
            metaList.forEach(meta -> stringBuilder.append(meta.toHtmlString()).append("\n"));
            linkList.forEach(link -> stringBuilder.append(link.toHtmlString()).append("\n"));

            if (this.isHasHtmlStyle) {
                stringBuilder.append(this.getStyleHtmlString());
            }

            stringBuilder.append("</head>\n");
            stringBuilder.append("<body>\n");
        } else {
            stringBuilder.append("<div class=\"exc-container\">\n");
            if (this.isHasHtmlStyle) {
                stringBuilder.append(this.getStyleHtmlString());
            }
        }

        elementList.stream().map(HtmlElement::toHtmlString).forEach(htmlString -> stringBuilder.append(htmlString).append("\n"));

        if (this.isHasHtmlContainer) {
            stringBuilder.append("</body>\n");
            stringBuilder.append("</html>\n");
        } else {
            stringBuilder.append("</div>\n");
        }
        return stringBuilder.toString();
    }
}
