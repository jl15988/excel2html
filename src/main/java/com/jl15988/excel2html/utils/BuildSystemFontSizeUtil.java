package com.jl15988.excel2html.utils;

import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.font.GlyphVector;
import java.awt.geom.Rectangle2D;
import java.util.HashMap;
import java.util.Map;

/**
 * 用于构建系统字体大小映射的工具
 *
 * @author Jalon
 * @since 2024/12/11 13:38
 **/
public class BuildSystemFontSizeUtil {

    public static void main(String[] args) {
        // 字体号
        int[] sizes = new int[]{6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72};
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();

        // 字体大小映射，字体名-{字体号-字体像素大小}
        Map<String, Map<Integer, Double>> map = new HashMap<>();

        // 获取所有可用的字体家族（Font Family）
        String[] fontFamilyNames = ge.getAvailableFontFamilyNames();
        for (String fontFamilyName : fontFamilyNames) {
            for (int size : sizes) {
                // 转换为AWT Font以进行字符宽度计算
                Font awtFont = new Font(fontFamilyName, Font.PLAIN, size);

                // 创建一个FontRenderContext，用于度量字体
                FontRenderContext frc = new FontRenderContext(null, true, true);
                String[] chas = new String[]{"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"};
                double charWidth = 0;
                for (String cha : chas) {
                    // 获取字符"0"的GlyphVector（用于度量）
                    GlyphVector gv = awtFont.createGlyphVector(frc, cha);
                    // 获取字符的边界框（Rectangle2D）
                    Rectangle2D rect = gv.getLogicalBounds();
                    // 计算字符宽度（以用户空间为单位）
                    double charWidthn = rect.getBounds().width;
                    if (charWidthn > charWidth) {
                        charWidth = charWidthn;
                    }
                }
//                System.out.println(charWidth);

                double dpi = 96.0;
                double charWidthInPixels = (charWidth * dpi) / 72.0; // 72点等于1英寸
                map.putIfAbsent(fontFamilyName, new HashMap<>());
                map.get(fontFamilyName).put(size, charWidthInPixels);
            }
        }

        // 构建映射put
        StringBuilder stringBuilder = new StringBuilder();
        map.forEach((s, integerDoubleMap) -> {
            stringBuilder.append("put(\"").append(s).append("\", ");
            stringBuilder.append("new HashMap<Integer, Double>() {{\n");
            integerDoubleMap.forEach((integer, aDouble) -> {
                stringBuilder.append("put(").append(integer).append(", ").append(aDouble).append(");\n");
            });
            stringBuilder.append("}}");
            stringBuilder.append(");\n");
        });
//        stringBuilder.append("}};");
        System.out.println(stringBuilder.toString());
    }
}
