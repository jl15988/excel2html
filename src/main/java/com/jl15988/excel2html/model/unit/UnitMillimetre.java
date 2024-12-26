package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.constant.UnitConstant;

/**
 * 毫米
 *
 * @author Jalon
 * @since 2024/12/10 10:55
 **/
public class UnitMillimetre extends Unit<UnitMillimetre> {

    public UnitMillimetre(double value, int dpi) {
        super(UnitMillimetre.class, value, dpi);
    }

    public UnitMillimetre(double value) {
        super(UnitMillimetre.class, value);
    }

    public static UnitMillimetre formPixel(UnitPixel unitPixel) {
        return Unit.to(unitPixel, UnitMillimetre.class);
    }

    public static UnitMillimetre formPoint(UnitPoint unitPoint) {
        return Unit.to(unitPoint, UnitMillimetre.class);
    }

    public static UnitMillimetre formMillimetre(UnitMillimetre unitMillimetre) {
        return Unit.to(unitMillimetre, UnitMillimetre.class);
    }

    public static UnitMillimetre formInch(UnitInch unitInch) {
        return Unit.to(unitInch, UnitMillimetre.class);
    }

    public static UnitMillimetre formEmu(UnitEmu unitEmu) {
        return Unit.to(unitEmu, UnitMillimetre.class);
    }

    /**
     * 转像素
     *
     * @param mm  毫米
     * @param dpi dpi
     * @return 像素
     */
    public static double toPixel(double mm, int dpi) {
        double point = toPoint(mm);
        return UnitPoint.toPixel(point, dpi);
    }

    /**
     * 转磅
     *
     * @param mm 毫米
     * @return 磅
     */
    public static double toPoint(double mm) {
        return 72 / 25.4 * mm;
    }

    /**
     * 转英寸
     *
     * @param mm  毫米
     * @param dpi dpi
     * @return 英寸
     */
    public static double toInch(double mm, int dpi) {
        return mm / 25.4;
    }

    /**
     * 转毫米
     *
     * @param mm 毫米
     * @return 毫米
     */
    public static double toMillimetre(double mm) {
        return mm;
    }

    /**
     * 转EMU
     *
     * @param mm 毫米
     * @return EMU
     */
    public static double toEmu(double mm) {
        return mm / UnitConstant.EMU_PER_CENTIMETER / 10;
    }
}
