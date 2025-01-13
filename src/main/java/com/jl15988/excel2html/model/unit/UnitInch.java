package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.constant.UnitConstant;

/**
 * 英寸
 *
 * @author Jalon
 * @since 2024/12/10 11:17
 **/
public class UnitInch extends Unit<UnitInch> {

    public UnitInch(double value, int dpi) {
        super(UnitInch.class, value, dpi);
    }

    public UnitInch(double value) {
        super(UnitInch.class, value);
    }

    public static UnitInch formPixel(UnitPixel unitPixel) {
        return Unit.to(unitPixel, UnitInch.class);
    }

    public static UnitInch formPoint(UnitPoint unitPoint) {
        return Unit.to(unitPoint, UnitInch.class);
    }

    public static UnitInch formMillimetre(UnitMillimetre unitMillimetre) {
        return Unit.to(unitMillimetre, UnitInch.class);
    }

    public static UnitInch formInch(UnitInch unitInch) {
        return Unit.to(unitInch, UnitInch.class);
    }

    public static UnitInch formEmu(UnitEmu unitEmu) {
        return Unit.to(unitEmu, UnitInch.class);
    }

    /**
     * 转像素
     *
     * @param inch 英寸
     * @param dpi  dpi
     * @return 像素
     */
    public static double toPixel(double inch, int dpi) {
        double point = toPoint(inch);
        return UnitPoint.toPixel(point, dpi);
    }

    /**
     * 转磅
     *
     * @param inch 英寸
     * @return 磅
     */
    public static double toPoint(double inch) {
        return inch * 72;
    }

    /**
     * 转英寸
     *
     * @param inch 英寸
     * @return 磅
     */
    public static double toInch(double inch) {
        return inch;
    }

    /**
     * 转毫米
     *
     * @param inch 英寸
     * @return 毫米
     */
    public static double toMillimetre(double inch) {
        return inch * 25.4;
    }

    /**
     * 转EMU
     *
     * @param inch 英寸
     * @return EMU
     */
    public static double toEmu(double inch) {
        return inch * UnitConstant.EMU_PER_INCH;
    }
}
