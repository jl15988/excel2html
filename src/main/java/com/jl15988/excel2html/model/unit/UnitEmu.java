package com.jl15988.excel2html.model.unit;

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

}
