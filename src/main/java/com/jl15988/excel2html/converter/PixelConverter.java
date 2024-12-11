package com.jl15988.excel2html.converter;

/**
 * 像素转换器
 *
 * @author Jalon
 * @since 2024/11/29 16:37
 **/
public class PixelConverter {

    /**
     * 像素转磅
     *
     * @param px  像素
     * @param dpi dpi
     * @return 磅
     */
    public static double toPoint(double px, int dpi) {
        return px * ((double) 72 / dpi);
    }

    /**
     * 像素转毫米
     *
     * @param px  像素
     * @param dpi dpi
     * @return 毫米
     */
    public static double toMM(double px, int dpi) {
        return (double) Math.round((px / (double) dpi * 25.4) * 100) / 100;
    }

    /**
     * 像素转英寸
     *
     * @param px  像素
     * @param dpi dpi
     * @return 英寸
     */
    public static double toInch(double px, int dpi) {
        double point = toPoint(px, dpi);
        return PointConverter.toInch(point);
    }

    /**
     * 像素转EMU
     *
     * @param px  像素
     * @param dpi dpi
     * @return EMU
     */
    public static double toEmu(double px, int dpi) {
        double mm = toMM(px, dpi);
        return MillimetreConverter.toEmu(mm);
    }
}
