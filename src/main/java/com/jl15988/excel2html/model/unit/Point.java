package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.converter.InchConverter;
import com.jl15988.excel2html.converter.MillimetreConverter;
import com.jl15988.excel2html.converter.PixelConverter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 磅
 *
 * @author Jalon
 * @since 2024/12/10 10:54
 **/
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Point {

    private double value;

    private int dpi = 96;

    private final String unit = "pt";

    public Point(double value) {
        this.value = value;
    }

    public Pixel toPixel() {
        return Pixel.formPoint(this);
    }

    public Millimetre toMillimetre() {
        return Millimetre.formPoint(this);
    }

    public Inch toInch() {
        return Inch.formPoint(this);
    }

    public static Point fromPixel(Pixel pixel) {
        return new Point(PixelConverter.toPoint(pixel.getValue(), pixel.getDpi()), pixel.getDpi());
    }

    public static Point fromMillimetre(Millimetre millimetre) {
        return new Point(MillimetreConverter.toPoint(millimetre.getValue()), millimetre.getDpi());
    }

    public static Point fromInch(Inch inch) {
        return new Point(InchConverter.toPoint(inch.getValue()), inch.getDpi());
    }

    public String toString() {
        return this.value + this.unit;
    }
}
