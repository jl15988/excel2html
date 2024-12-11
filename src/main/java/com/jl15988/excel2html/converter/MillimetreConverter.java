package com.jl15988.excel2html.converter;

/**
 * @author Jalon
 * @since 2024/12/10 11:22
 **/
public class MillimetreConverter {

    /**
     * 毫米转像素
     *
     * @param mm  毫米
     * @param dpi dpi
     * @return 像素
     */
    public static double toPx(double mm, int dpi) {
        double point = toPoint(mm);
        return PointConverter.toPx(point, dpi);
    }

    /**
     * 毫米转磅
     *
     * @param mm 毫米
     * @return 磅
     */
    public static double toPoint(double mm) {
        return 72 / 25.4 * mm;
    }

    /**
     * 毫米转英寸
     *
     * @param mm  毫米
     * @param dpi dpi
     * @return 英寸
     */
    public static double toInch(double mm, int dpi) {
        return mm / 25.4;
    }

    /**
     * 毫米转EMU
     *
     * @param mm 毫米
     * @return EMU
     */
    public static double toEmu(double mm) {
        return mm / UnitConstant.EMU_PER_CENTIMETER / 10;
    }
}
