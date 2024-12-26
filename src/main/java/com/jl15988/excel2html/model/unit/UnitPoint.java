package com.jl15988.excel2html.model.unit;

/**
 * 磅
 *
 * @author Jalon
 * @since 2024/12/10 10:54
 **/
public class UnitPoint extends Unit<UnitPoint> {

    public UnitPoint(double value, int dpi) {
        super(UnitPoint.class, value, dpi);
    }

    public UnitPoint(double value) {
        super(UnitPoint.class, value);
    }

    public static UnitPoint formPixel(UnitPixel unitPixel) {
        return Unit.to(unitPixel, UnitPoint.class);
    }

    public static UnitPoint formPoint(UnitPoint unitPoint) {
        return Unit.to(unitPoint, UnitPoint.class);
    }

    public static UnitPoint formMillimetre(UnitMillimetre unitMillimetre) {
        return Unit.to(unitMillimetre, UnitPoint.class);
    }

    public static UnitPoint formInch(UnitInch unitInch) {
        return Unit.to(unitInch, UnitPoint.class);
    }

    public static UnitPoint formEmu(UnitEmu unitEmu) {
        return Unit.to(unitEmu, UnitPoint.class);
    }

    /**
     * 转像素
     *
     * @param points 磅
     * @return 像素
     */
    public static double toPixel(double points, int dpi) {
        // 英寸
        double inch = 72;
        // 1磅=1/72英寸，而1英寸所含有DPI个像素
        // 近似转换为像素,96dpi / 72英寸
        return points * (((double) dpi) / inch);
    }

    /**
     * 转磅
     *
     * @param points 磅
     * @return 磅
     */
    public static double toPoint(double points) {
        return points;
    }

    /**
     * 转毫米
     *
     * @param points 磅
     * @return 毫米
     */
    public static double toMillimetre(double points, int dpi) {
        double px = toPixel(points, dpi);
        return UnitPixel.toMillimetre(px, dpi);
    }

    /**
     * 转英寸
     *
     * @param points 磅
     * @return 英寸
     */
    public static double toInch(double points) {
        return points / 72;
    }

    /**
     * 转EMU
     *
     * @param points 磅
     * @return EMU
     */
    public static double toEmu(double points) {
        double inch = toInch(points);
        return UnitInch.toEmu(inch);
    }
}
