package com.jl15988.excel2html.converter;

import org.apache.poi.util.Units;

import java.text.DecimalFormat;

/**
 * 像素转换器
 *
 * @author Jalon
 * @since 2024/11/29 16:37
 **/
public class PixelConverter {

    public static float cellPixelsToPx(double cellPixels) {
        return (float) (cellPixels * 1.2);
    }

    public static String cellPixelsToPxString(double cellPixels) {
        DecimalFormat df = new DecimalFormat("0.00");
        return df.format(cellPixelsToPx(cellPixels)) + "px";
    }

    /**
     * 磅转像素
     *
     * @param points 磅
     * @return 像素
     */
    public static float pointsToPx(double points) {
        // DPI
        double dpi = 96;
        // 英寸
        double inch = 72;
        // 变量
        double var = 1.03;
        // 1磅=1/72英寸，而1英寸所含有DPI个像素
        // 近似转换为像素,96dpi / 72英寸
        double pxD = (double) points * (dpi / inch) * 1.03;
        return (float) pxD;
//        return Units.pointsToPixel(points);
    }

    public static String pointsToPxString(double points) {
        DecimalFormat df = new DecimalFormat("0.00");
        return df.format(pointsToPx(points)) + "px";
    }

    public static double emusToPx(long emus) {
        double points = Units.toPoints(emus);
        return pointsToPx(points);
    }

    public static String emusToPxString(long emus) {
        DecimalFormat df = new DecimalFormat("0.00");
        return df.format(emusToPx(emus)) + "px";
    }
}