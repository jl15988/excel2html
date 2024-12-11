package com.jl15988.excel2html.converter;

/**
 * 英寸转换器
 *
 * @author Jalon
 * @since 2024/12/10 11:34
 **/
public class InchConverter {

    /**
     * 英寸转像素
     *
     * @param inch 英寸
     * @param dpi  dpi
     * @return 像素
     */
    public static double toPx(double inch, int dpi) {
        double point = toPoint(inch);
        return PointConverter.toPx(point, dpi);
    }

    /**
     * 英寸转磅
     *
     * @param inch 英寸
     * @return 磅
     */
    public static double toPoint(double inch) {
        return inch * 72;
    }

    /**
     * 英寸转毫米
     *
     * @param inch 英寸
     * @return 毫米
     */
    public static double toMM(double inch) {
        return inch * 25.4;
    }

    /**
     * 英寸转EMU
     *
     * @param inch 英寸
     * @return EMU
     */
    public static double toEmu(double inch) {
        return inch / UnitConstant.EMU_PER_INCH;
    }
}
