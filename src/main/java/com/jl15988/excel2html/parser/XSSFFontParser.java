package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.converter.ColorConverter;
import com.jl15988.excel2html.converter.PixelConverter;
import com.jl15988.excel2html.model.style.FontICssStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Objects;

/**
 * @author Jalon
 * @since 2024/11/29 17:02
 **/
public class XSSFFontParser {

    /**
     * 解析文字为css样式
     *
     * @param xssfFont 文字
     */
    public static FontICssStyle parserXSSFFontToStyleMap(XSSFFont xssfFont) {
        FontICssStyle fontCssStyle = new FontICssStyle();
        if (Objects.isNull(xssfFont)) {
            return fontCssStyle;
        }
        boolean isBold = xssfFont.getBold();
        boolean isItalic = xssfFont.getItalic();
        boolean isStrikeout = xssfFont.getStrikeout();
        boolean isUnderline = xssfFont.getUnderline() != Font.U_NONE;
        String fontName = xssfFont.getFontName();
        short fontHeightInPoints = xssfFont.getFontHeightInPoints();
        XSSFColor xssfColor = xssfFont.getXSSFColor();
        String rgba = ColorConverter.xSSFColorToRGBAString(xssfColor);

        fontCssStyle.set(isBold, "font-weight", "bold")
                .set(isItalic, "font-style", "italic")
                .set(isStrikeout, "text-decoration", "line-through")
                .set(isUnderline, "text-decoration", "underline")
                .set("font-family", fontName)
                .set("font-size", PixelConverter.pointsToPxString(fontHeightInPoints))
                .setIfExists("color", rgba);

        return fontCssStyle;
    }
}
