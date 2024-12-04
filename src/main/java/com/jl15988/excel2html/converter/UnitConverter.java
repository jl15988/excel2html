package com.jl15988.excel2html.converter;

import org.apache.poi.util.Units;

import java.text.DecimalFormat;

/**
 * 单位转换器
 *
 * @author Jalon
 * @since 2024/12/3 14:03
 **/
public class UnitConverter {

    private boolean usePx = false;

    public static UnitConverter convert() {
        return new UnitConverter();
    }

    public UnitConverter usePx() {
        this.usePx = true;
        return this;
    }

    public UnitConverter setUsePx(boolean usePx) {
        this.usePx = usePx;
        return this;
    }

    public boolean getUsePx() {
        return usePx;
    }

    public String getUnit() {
        return usePx ? "px" : "pt";
    }

    public String formatVal(double val) {
        DecimalFormat df = new DecimalFormat("0.00");
        return df.format(val);
    }

    public double convertCellPixels(double cellPixels) {
        if (usePx) {
            return PixelConverter.cellPixelsToPx(cellPixels);
        }
        // todo 可能不准
        return Units.pixelToPoints(cellPixels);
    }

    public String convertCellPixelsString(double cellPixels) {
        return formatVal(convertCellPixels(cellPixels)) + getUnit();
    }

    public double convertPoints(double points) {
        if (usePx) {
            return PixelConverter.pointsToPx(points);
        }
        return points;
    }

    public String convertPointsString(double points) {
        return formatVal(convertPoints(points)) + getUnit();
    }

    public double convertEmus(long emus) {
        if (usePx) {
            return PixelConverter.emusToPx(emus);
        }
        return Units.toPoints(emus);
    }

    public String convertEmusString(long emus) {
        return formatVal(convertEmus(emus)) + getUnit();
    }

}
