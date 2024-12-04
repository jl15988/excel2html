package com.jl15988.excel2html.html;

import java.util.List;
import java.util.Map;

public interface IHtmlElement<T> {

    /**
     * 获取唯一标识
     */
    String getUID();

    /**
     * 获取id
     */
    String getId();

    /**
     * 设置id
     *
     * @param id id
     */
    T setId(String id);

    /**
     * 获取标签名称
     */
    String getTagName();

    /**
     * 判断是否某个标签
     *
     * @param tagName 标签名称
     */
    default boolean isTag(String tagName) {
        return getTagName().equals(tagName);
    }

    /**
     * 获取class列表
     */
    List<String> getClassList();

    /**
     * 添加class
     *
     * @param className class名称
     */
    T addClass(String className);

    /**
     * 添加class
     *
     * @param classList class名称集合
     */
    T addClasses(List<String> classList);

    /**
     * 添加class
     *
     * @param classes class名称
     */
    T addClasses(String... classes);

    /**
     * 移除class名
     *
     * @param className class名称
     */
    T removeClass(String className);

    /**
     * 清除所有class
     */
    T clearClass();

    /**
     * 判断是否包含某个class
     *
     * @param className class名
     */
    default boolean hasClass(String className) {
        return getClassList().contains(className);
    }

    /**
     * 获取样式
     */
    Map<String, Object> getStyleMap();

    /**
     * 设置样式
     *
     * @param styleMap 样式
     */
    T setStyleMap(Map<String, Object> styleMap);

    /**
     * 获取样式值
     *
     * @param styleName 样式名
     */
    Object getStyle(String styleName);

    /**
     * 添加样式
     *
     * @param styleName 样式名
     * @param value     值
     */
    T addStyle(String styleName, Object value);

    /**
     * 添加样式
     *
     * @param cssStyle css样式
     */
    T addStyle(ICssStyle<?> cssStyle);

    /**
     * 移除某个样式
     *
     * @param styleName 样式名
     */
    T removeStyle(String styleName);

    /**
     * 清除所有样式
     */
    T clearStyle();

    /**
     * 判断是否有某个样式
     *
     * @param styleName 样式名
     */
    default boolean hasStyle(String styleName) {
        return getStyleMap().containsKey(styleName);
    }

    /**
     * 判断是否有某个样式，包括值
     *
     * @param styleName 样式名
     * @param value     值
     */
    default boolean hasStyle(String styleName, Object value) {
        if (!getStyleMap().containsKey(styleName)) {
            return false;
        }
        return getStyleMap().get(styleName).equals(value);
    }

    /**
     * 获取属性列表
     */
    Map<String, String> getAttributeMap();

    /**
     * 获取属性值
     *
     * @param name 属性名
     */
    String getAttribute(String name);

    /**
     * 添加属性
     *
     * @param name  属性名
     * @param value 属性值
     */
    T addAttribute(String name, String value);

    /**
     * 移除某个属性
     *
     * @param name 属性名
     */
    T removeAttribute(String name);

    /**
     * 清除所有属性
     */
    T clearAttributes();

    /**
     * 判断是否包含某个属性
     *
     * @param name 属性名
     */
    default boolean hasAttribute(String name) {
        return getAttributeMap().containsKey(name);
    }

    /**
     * 判断是否包含某个属性，包括值
     *
     * @param name  属性名
     * @param value 属性值
     */
    default boolean hasAttribute(String name, String value) {
        if (!getAttributeMap().containsKey(name)) {
            return false;
        }
        return getAttributeMap().get(name).equals(value);
    }

    /**
     * 获取所有子元素
     */
    List<IHtmlElement<?>> getChildrenElementList();

    /**
     * 添加子元素
     *
     * @param child 子元素
     */
    T addChildElement(IHtmlElement<?> child);

    /**
     * 添加子元素
     *
     * @param index 下标
     * @param child 子元素
     */
    T addChildElement(int index, IHtmlElement<?> child);

    /**
     * 移除某个子元素
     *
     * @param child 子元素
     */
    T removeChild(IHtmlElement<?> child);

    /**
     * 移除某个子元素
     *
     * @param index 下标
     */
    T removeChild(int index);

    /**
     * 清除所有子元素
     */
    T clearChildren();

    /**
     * 判断是否包含某个子元素，只包含一代子元素
     *
     * @param child 子元素
     */
    default boolean hasChildrenElement(IHtmlElement<?> child) {
        return getChildrenElementList().stream().anyMatch(IHtmlElement -> IHtmlElement.getUID().equals(child.getUID()));
    }

    /**
     * 设置内容，将会覆盖所有子元素
     *
     * @param content 内容
     */
    T setContent(String content);

    /**
     * 获取内容
     */
    String getContent();

    /**
     * 转为html字符
     */
    String toHtmlString();

    default boolean equals(IHtmlElement<?> other) {
        return getUID().equals(other.getUID());
    }
}
