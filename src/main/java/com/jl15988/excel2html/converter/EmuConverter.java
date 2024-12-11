package com.jl15988.excel2html.converter;

/**
 * emu 转换器
 *
 * @author Jalon
 * @since 2024/12/11 20:59
 **/
public class EmuConverter {

    /**
     * EMU转像素
     *
     * @param emu EMU
     * @param dpi dpi
     * @return 像素
     */
    public static double toPx(double emu, int dpi) {
        double point = toPoint(emu);
        return PointConverter.toPx(point, dpi);
    }

    /**
     * EMU转磅
     *
     * @param emu EMU
     * @return 磅
     */
    public static double toPoint(double emu) {
        double inch = toInch(emu);
        return InchConverter.toPoint(inch);
    }

    /**
     * EMU转英寸
     *
     * @param emu EMU
     * @return 磅
     */
    public static double toInch(double emu) {
        return emu * UnitConstant.EMU_PER_INCH;
    }

    /**
     * EMU转毫米
     *
     * @param emu EMU
     * @return 毫米
     */
    public static double toMM(double emu) {
        return emu * UnitConstant.EMU_PER_CENTIMETER * 10;
    }
}
