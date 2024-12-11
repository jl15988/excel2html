package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.converter.InchConverter;
import com.jl15988.excel2html.converter.MillimetreConverter;
import com.jl15988.excel2html.converter.PixelConverter;
import com.jl15988.excel2html.converter.PointConverter;
import com.jl15988.excel2html.converter.UnitConstant;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * EMU
 *
 * @author Jalon
 * @since 2024/12/11 21:08
 **/
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Emu {

    private double value;

    private int dpi = UnitConstant.DEFAULT_DPI;

    public Pixel toPixel() {
        return Pixel.formEmu(this);
    }

    public Point toPoint() {
        return Point.formEmu(this);
    }

    public Millimetre toMillimetre() {
        return Millimetre.formEmu(this);
    }

    public Inch toInch() {
        return Inch.formEmu(this);
    }

    public static Emu formPixel(Pixel pixel) {
        return new Emu(PixelConverter.toEmu(pixel.getValue(), pixel.getDpi()), pixel.getDpi());
    }

    public static Emu formPoint(Point point) {
        return new Emu(PointConverter.toEmu(point.getValue()), point.getDpi());
    }

    public static Emu formMillimetre(Millimetre millimetre) {
        return new Emu(MillimetreConverter.toEmu(millimetre.getValue()), millimetre.getDpi());
    }

    public static Emu formInch(Inch inch) {
        return new Emu(InchConverter.toEmu(inch.getValue()), inch.getDpi());
    }

    public String toString() {
        return String.valueOf(this.value);
    }
}
