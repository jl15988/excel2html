package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.converter.EmuConverter;
import com.jl15988.excel2html.converter.MillimetreConverter;
import com.jl15988.excel2html.converter.PixelConverter;
import com.jl15988.excel2html.converter.PointConverter;
import com.jl15988.excel2html.converter.UnitConstant;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 英寸
 *
 * @author Jalon
 * @since 2024/12/10 11:17
 **/
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Inch {

    private double value;

    private int dpi = UnitConstant.DEFAULT_DPI;

    private final String unit = UnitConstant.INCH_UNIT;

    public Inch(double value) {
        this.value = value;
    }

    public Pixel toPixel() {
        return Pixel.formInch(this);
    }

    public Point toPoint() {
        return Point.fromInch(this);
    }

    public Millimetre toMillimetre() {
        return Millimetre.formInch(this);
    }

    public Emu toEmu() {
        return Emu.formInch(this);
    }

    public static Inch formPixel(Pixel pixel) {
        return new Inch(PixelConverter.toInch(pixel.getValue(), pixel.getDpi()), pixel.getDpi());
    }

    public static Inch formPoint(Point point) {
        return new Inch(PointConverter.toInch(point.getValue()), point.getDpi());
    }

    public static Inch formMillimetre(Millimetre millimetre) {
        return new Inch(MillimetreConverter.toInch(millimetre.getValue(), millimetre.getDpi()), millimetre.getDpi());
    }

    public static Inch formEmu(Emu emu) {
        return new Inch(EmuConverter.toInch(emu.getValue()), emu.getDpi());
    }

    public String toString() {
        return this.value + this.unit;
    }
}
