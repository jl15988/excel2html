package com.jl15988.excel2html.model.style;

import com.jl15988.excel2html.html.CssStyle;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/3 15:55
 **/
public class CommonCss {

    public Map<String, CssStyle> commonCss = new HashMap<String, CssStyle>();

    public static final String ALIGN_HORIZONTAL_GENERAL_NUMERIC = "align--row--genera--numeric";
    public static final String ALIGN_HORIZONTAL_GENERAL_STRING = "align--row--general--string";
    public static final String ALIGN_HORIZONTAL_GENERAL_BOOLEAN = "align--row--general--boolean";
    public static final String ALIGN_HORIZONTAL_LEFT = "align--row--left";
    public static final String ALIGN_HORIZONTAL_CENTER = "align--row--center";
    public static final String ALIGN_HORIZONTAL_RIGHT = "align--row--right";
    public static final String ALIGN_HORIZONTAL_FILL = "align--row--fill";
    public static final String ALIGN_HORIZONTAL_JUSTIFY = "align--row--justify";
    public static final String ALIGN_HORIZONTAL_CENTER_SELECTION = "align--row--center_selection";
    public static final String ALIGN_HORIZONTAL_DISTRIBUTED = "align--row--distributed";
    public static final String ALIGN_VERTICAL_TOP = "align--col--top";
    public static final String ALIGN_VERTICAL_CENTER = "align--col--center";
    public static final String ALIGN_VERTICAL_BOTTOM = "align--col--bottom";
    public static final String ALIGN_VERTICAL_JUSTIFY = "align--col--justify";
    public static final String ALIGN_VERTICAL_DISTRIBUTED = "align--col--distributed";

    public CommonCss() {
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_GENERAL_NUMERIC,
                new CssStyle()
                        .set("text-align", "right")
                        .set("justify-content", "flex-end")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_GENERAL_STRING,
                new CssStyle()
                        .set("text-align", "left")
                        .set("justify-content", "flex-start")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_GENERAL_BOOLEAN,
                new CssStyle()
                        .set("text-align", "center")
                        .set("justify-content", "center")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_LEFT,
                new CssStyle()
                        .set("text-align", "left")
                        .set("justify-content", "flex-start")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_CENTER,
                new CssStyle()
                        .set("text-align", "center")
                        .set("justify-content", "center")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_RIGHT,
                new CssStyle()
                        .set("text-align", "right")
                        .set("justify-content", "flex-end")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_FILL,
                new CssStyle()
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_JUSTIFY,
                new CssStyle()
                        .set("text-align", "justify")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_CENTER_SELECTION,
                new CssStyle()
                        .set("text-align", "center")
        );
        commonCss.put(
                CommonCss.ALIGN_HORIZONTAL_DISTRIBUTED,
                new CssStyle()
                        .set("text-align", "justify")
                        .set("text-align-last", "justify")
        );
        commonCss.put(
                CommonCss.ALIGN_VERTICAL_TOP,
                new CssStyle()
                        .set("vertical-align", "baseline")
                        .set("align-items", "flex-start")
        );
        commonCss.put(
                CommonCss.ALIGN_VERTICAL_CENTER,
                new CssStyle()
                        .set("vertical-align", "middle")
                        .set("align-items", "center")
        );
        commonCss.put(
                CommonCss.ALIGN_VERTICAL_BOTTOM,
                new CssStyle()
                        .set("vertical-align", "bottom")
                        .set("align-items", "flex-end")
        );
        commonCss.put(
                CommonCss.ALIGN_VERTICAL_JUSTIFY,
                new CssStyle()
        );
        commonCss.put(
                CommonCss.ALIGN_VERTICAL_DISTRIBUTED,
                new CssStyle()
        );
    }

    public String toHtmlString() {
        StringBuilder stringBuilder = new StringBuilder();
        for (Map.Entry<String, CssStyle> entry : commonCss.entrySet()) {
            String key = entry.getKey();
            CssStyle cssStyle = entry.getValue();
            stringBuilder.append(".").append(key).append(" {").append(cssStyle.toHtmlString()).append("}");
        }
        return stringBuilder.toString();
    }
}
