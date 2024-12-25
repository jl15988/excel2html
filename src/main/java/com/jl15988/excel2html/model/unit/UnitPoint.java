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
}
