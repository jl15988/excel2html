package com.jl15988.excel2html.converter;

/**
 * 单位常量
 *
 * @author Jalon
 * @since 2024/12/11 20:39
 **/
public class UnitConstant {
    // English Metric Unit，EMU是用于度量单位转换的一个系统，主要用于图形对象的尺寸计算
    // 每厘米EMU数量
    public static final int EMU_PER_CENTIMETER = 360000;
    // 每英寸EMU数量
    public static final int EMU_PER_INCH = 914400;

    // 默认DPI
    public static final int DEFAULT_DPI = 96;

    // 默认像素单位
    public static final String PIXEL_UNIT = "px";
    // 默认磅单位
    public static final String POINT_UNIT = "pt";
    // 默认毫米单位
    public static final String MILLIMETRE_UNIT = "mm";
    // 默认厘米单位
    public static final String CENTIMETER_UNIT = "cm";
    // 默认英寸单位
    public static final String INCH_UNIT = "in";
    // 默认EMU单位
    public static final String ENUM_UNIT = "emu";
}
