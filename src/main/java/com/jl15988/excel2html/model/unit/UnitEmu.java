package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.constant.UnitConstant;

/**
 * EMU
 *
 * @author Jalon
 * @since 2024/12/11 21:08
 **/
public class UnitEmu extends Unit<UnitEmu> {

    public UnitEmu(double value, int dpi) {
        super(UnitEmu.class, value, dpi);
    }

    public UnitEmu(double value) {
        super(UnitEmu.class, value);
    }

    public static UnitEmu formPixel(UnitPixel unitPixel) {
        return Unit.to(unitPixel, UnitEmu.class);
    }

    public static UnitEmu formPoint(UnitPoint unitPoint) {
        return Unit.to(unitPoint, UnitEmu.class);
    }

    public static UnitEmu formMillimetre(UnitMillimetre unitMillimetre) {
        return Unit.to(unitMillimetre, UnitEmu.class);
    }

    public static UnitEmu formInch(UnitInch unitInch) {
        return Unit.to(unitInch, UnitEmu.class);
    }

    public static UnitEmu formEmu(UnitEmu unitEmu) {
        return Unit.to(unitEmu, UnitEmu.class);
    }

    /**
     * 转像素
     *
     * @param emu EMU
     * @param dpi dpi
     * @return 像素
     */
    public static double toPixel(double emu, int dpi) {
        double point = toPoint(emu);
        return UnitPoint.toPixel(point, dpi);
    }

    /**
     * 转磅
     *
     * @param emu EMU
     * @return 磅
     */
    public static double toPoint(double emu) {
        double inch = toInch(emu);
        return UnitInch.toPoint(inch);
    }

    /**
     * 转英寸
     *
     * @param emu EMU
     * @return 磅
     */
    public static double toInch(double emu) {
        return emu * UnitConstant.EMU_PER_INCH;
    }

    /**
     * 转毫米
     *
     * @param emu EMU
     * @return 毫米
     */
    public static double toMillimetre(double emu) {
        return emu * UnitConstant.EMU_PER_CENTIMETER * 10;
    }

    /**
     * 转EMU
     *
     * @param emu EMU
     * @return EMU
     */
    public static double toEmu(double emu) {
        return emu;
    }

}
