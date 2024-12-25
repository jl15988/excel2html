package com.jl15988.excel2html.model.unit;

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
}
