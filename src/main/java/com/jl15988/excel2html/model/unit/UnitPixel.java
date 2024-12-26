package com.jl15988.excel2html.model.unit;

/**
 * 像素
 *
 * @author Jalon
 * @since 2024/12/10 10:54
 **/
public class UnitPixel extends Unit<UnitPixel> {

    public UnitPixel(double value, int dpi) {
        super(UnitPixel.class, value, dpi);
    }

    public UnitPixel(double value) {
        super(UnitPixel.class, value);
    }

    public static UnitPixel formPixel(UnitPixel unitPixel) {
        return Unit.to(unitPixel, UnitPixel.class);
    }

    public static UnitPixel formPoint(UnitPoint unitPoint) {
        return Unit.to(unitPoint, UnitPixel.class);
    }

    public static UnitPixel formMillimetre(UnitMillimetre unitMillimetre) {
        return Unit.to(unitMillimetre, UnitPixel.class);
    }

    public static UnitPixel formInch(UnitInch unitInch) {
        return Unit.to(unitInch, UnitPixel.class);
    }

    public static UnitPixel formEmu(UnitEmu unitEmu) {
        return Unit.to(unitEmu, UnitPixel.class);
    }

    /**
     * 转像素
     *
     * @param px 像素
     * @return 像素
     */
    public static double toPixel(double px) {
        return px;
    }

    /**
     * 转磅
     *
     * @param px  像素
     * @param dpi dpi
     * @return 磅
     */
    public static double toPoint(double px, int dpi) {
        return px * ((double) 72 / dpi);
    }

    /**
     * 转毫米
     *
     * @param px  像素
     * @param dpi dpi
     * @return 毫米
     */
    public static double toMillimetre(double px, int dpi) {
        return (double) Math.round((px / (double) dpi * 25.4) * 100) / 100;
    }

    /**
     * 转英寸
     *
     * @param px  像素
     * @param dpi dpi
     * @return 英寸
     */
    public static double toInch(double px, int dpi) {
        double point = toPoint(px, dpi);
        return UnitPoint.toInch(point);
    }

    /**
     * 转EMU
     *
     * @param px  像素
     * @param dpi dpi
     * @return EMU
     */
    public static double toEmu(double px, int dpi) {
        double mm = toMillimetre(px, dpi);
        return UnitMillimetre.toEmu(mm);
    }
}
