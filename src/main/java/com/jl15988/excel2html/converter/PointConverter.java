package com.jl15988.excel2html.converter;

/**
 * 磅转换器
 *
 * @author Jalon
 * @since 2024/12/10 11:04
 **/
public class PointConverter {

    /**
     * 磅转像素
     *
     * @param points 磅
     * @return 像素
     */
    public static double toPx(double points, int dpi) {
        // 英寸
        double inch = 72;
        // 1磅=1/72英寸，而1英寸所含有DPI个像素
        // 近似转换为像素,96dpi / 72英寸
        return points * (((double) dpi) / inch);
    }

    /**
     * 磅转毫米
     *
     * @param points 磅
     * @return 毫米
     */
    public static double toMM(double points, int dpi) {
        double px = toPx(points, dpi);
        return PixelConverter.toMM(px, dpi);
    }

    /**
     * 磅转英寸
     *
     * @param points 磅
     * @return 英寸
     */
    public static double toInch(double points) {
        return points / 72;
    }

    /**
     * 磅转EMU
     *
     * @param points 磅
     * @return EMU
     */
    public static double toEmu(double points) {
        double inch = toInch(points);
        return InchConverter.toEmu(inch);
    }
}
