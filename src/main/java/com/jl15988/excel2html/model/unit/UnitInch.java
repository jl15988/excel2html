package com.jl15988.excel2html.model.unit;

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
}
