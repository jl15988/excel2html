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
}
