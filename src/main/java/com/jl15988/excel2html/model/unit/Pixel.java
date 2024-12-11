package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.converter.InchConverter;
import com.jl15988.excel2html.converter.MillimetreConverter;
import com.jl15988.excel2html.converter.PointConverter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 像素
 *
 * @author Jalon
 * @since 2024/12/10 10:54
 **/
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Pixel {

    private double value;

    private int dpi = 96;

    private final String unit = "px";

    public Pixel(double value) {
        this.value = value;
    }

    public Point toPoint() {
        return Point.fromPixel(this);
    }

    public Millimetre toMillimetre() {
        return Millimetre.formPixel(this);
    }

    public Inch toInch() {
        return Inch.formPixel(this);
    }

    public static Pixel formPoint(Point point) {
        return new Pixel(PointConverter.toPx(point.getValue(), point.getDpi()), point.getDpi());
    }

    public static Pixel formMillimetre(Millimetre millimetre) {
        return new Pixel(MillimetreConverter.toPx(millimetre.getValue(), millimetre.getDpi()), millimetre.getDpi());
    }

    public static Pixel formInch(Inch inch) {
        return new Pixel(InchConverter.toPx(inch.getValue(), inch.getDpi()), inch.getDpi());
    }

    public String toString() {
        return this.value + this.unit;
    }
}
