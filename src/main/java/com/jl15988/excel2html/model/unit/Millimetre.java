package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.converter.EmuConverter;
import com.jl15988.excel2html.converter.InchConverter;
import com.jl15988.excel2html.converter.PixelConverter;
import com.jl15988.excel2html.converter.PointConverter;
import com.jl15988.excel2html.converter.UnitConstant;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 毫米
 *
 * @author Jalon
 * @since 2024/12/10 10:55
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class Millimetre {

    private double value;

    private int dpi = UnitConstant.DEFAULT_DPI;

    private final String unit = UnitConstant.MILLIMETRE_UNIT;

    public Millimetre(double value) {
        this.value = value;
    }

    public Pixel toPixel() {
        return Pixel.formMillimetre(this);
    }

    public Point toPoint() {
        return Point.fromMillimetre(this);
    }

    public Inch toInch() {
        return Inch.formMillimetre(this);
    }

    public Emu toEmu() {
        return Emu.formMillimetre(this);
    }

    public static Millimetre formPixel(Pixel pixel) {
        return new Millimetre(PixelConverter.toMM(pixel.getValue(), pixel.getDpi()), pixel.getDpi());
    }

    public static Millimetre formPoint(Point point) {
        return new Millimetre(PointConverter.toMM(point.getValue(), point.getDpi()));
    }

    public static Millimetre formInch(Inch inch) {
        return new Millimetre(InchConverter.toMM(inch.getValue()), inch.getDpi());
    }

    public static Millimetre formEmu(Emu emu) {
        return new Millimetre(EmuConverter.toMM(emu.getValue()), emu.getDpi());
    }

    public String toString() {
        return this.value + this.unit;
    }
}
