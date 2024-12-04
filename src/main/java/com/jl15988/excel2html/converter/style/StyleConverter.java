package com.jl15988.excel2html.converter.style;

import com.jl15988.excel2html.utils.CssUtil;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * 样式转换器
 *
 * @author Jalon
 * @since 2024/12/2 10:28
 **/
public class StyleConverter {

    /**
     * style 分组，tagUid-styleMap<styleName, styleValue>
     *
     * @param tagStyleMap tag-style 集
     */
    public static List<StyleGroup> tagStyleToStyleGroup(Map<String, Map<String, Object>> tagStyleMap) {
        Set<String> tagUIDs = tagStyleMap.keySet();

        Map<String, List<String>> tagStylesMap = new HashMap<>();

        for (String tagUID : tagUIDs) {
            Map<String, Object> styleMap = tagStyleMap.get(tagUID);

            List<String> styles = new ArrayList<>();
            styleMap.forEach((name, value) -> {
                String styleVal = name + ":" + String.valueOf(value);
                styles.add(styleVal);
            });

            tagStylesMap.put(tagUID, styles);
        }

        // style 分组
        List<StyleGroup> styleGroups = new ArrayList<>();

        // 用于缓存生成的 style-id
        Map<String, String> stylesMap = new HashMap<>();
        AtomicInteger i = new AtomicInteger();
        tagStylesMap.forEach((tagUID, styles) -> {
            String stylesVal = styles.stream().sorted(String::compareTo).collect(Collectors.joining(","));

            // 构建样式分组
            StyleGroup styleGroup = new StyleGroup();
            styleGroup.setTagUid(tagUID);
            // 解析样式为 map
            Map<String, Object> styleMap = new HashMap<>();
            for (String style : styles) {
                String[] split = style.split(":");
                styleMap.put(split[0], split[1]);
            }
            styleGroup.setStyleMap(styleMap);

            String styleUid;
            if (stylesMap.containsKey(stylesVal)) {
                // 已存在样式，则使用存在的 id
                styleUid = stylesMap.get(stylesVal);
            } else {
                // 生成新的 id
                styleUid = CssUtil.randomName(5, String.valueOf(i.intValue()));
                stylesMap.put(stylesVal, styleUid);
            }
            styleGroup.setStyleUid(styleUid);
            styleGroups.add(styleGroup);

            i.incrementAndGet();
        });

        return styleGroups;
    }

    /**
     * style group 转 html
     *
     * @param styleGroups style 组
     */
    public static StyleGroupHtml styleGroupToHtmlString(List<StyleGroup> styleGroups) {
        // 标签-样式id
        Map<String, List<String>> tagStyleUidMap = new HashMap<>();
        // 样式id-样式
        Map<String, Map<String, Object>> uidStyleMap = new HashMap<>();

        for (StyleGroup styleGroup : styleGroups) {
            // 添加标签-样式映射
            tagStyleUidMap.putIfAbsent(styleGroup.getTagUid(), new ArrayList<>());
            tagStyleUidMap.get(styleGroup.getTagUid()).add(styleGroup.getStyleUid());

            // 添加样式id-样式映射
            if (!uidStyleMap.containsKey(styleGroup.getStyleUid())) {
                uidStyleMap.put(styleGroup.getStyleUid(), styleGroup.getStyleMap());
            }
        }

        // 构建css样式
        StringBuilder stringBuilder = new StringBuilder();
        uidStyleMap.forEach((uid, styleMap) -> {
            stringBuilder.append(" .").append(uid).append(" {");
            styleMap.forEach((key, value) -> {
                stringBuilder.append(" ").append(key).append(": ").append(value).append("; ");
            });
            stringBuilder.append("}");
        });

        StyleGroupHtml styleGroupHtml = new StyleGroupHtml();
        styleGroupHtml.setStyleContent(stringBuilder.toString());
        styleGroupHtml.setTagStyleUidMap(tagStyleUidMap);
        return styleGroupHtml;
    }

    /**
     * 将 tag-style 查重，转换为 css 和 css class
     *
     * @param tagStyleMap tag-style 集
     */
    public static StyleGroupHtml tagStyleToHtmlString(Map<String, Map<String, Object>> tagStyleMap) {
        List<StyleGroup> styleGroups = tagStyleToStyleGroup(tagStyleMap);
        return styleGroupToHtmlString(styleGroups);
    }
}
