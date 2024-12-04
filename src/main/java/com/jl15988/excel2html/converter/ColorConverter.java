package com.jl15988.excel2html.converter;

import java.util.HashMap;
import java.util.Map;

/**
 * 颜色转换器
 *
 * @author Jalon
 * @since 2024/11/29 16:41
 **/
public class ColorConverter {

    public static final Map<String, String> hexAlphaMap = new HashMap<String, String>() {{
        put("FF", "1");
        put("E6", "0.90");
        put("E3", "0.89");
        put("E0", "0.88");
        put("DE", "0.87");
        put("DB", "0.86");
        put("D9", "0.85");
        put("D6", "0.84");
        put("D4", "0.83");
        put("D1", "0.82");
        put("CF", "0.81");
        put("CC", "0.80");
        put("C9", "0.79");
        put("C7", "0.78");
        put("C4", "0.77");
        put("C2", "0.76");
        put("BF", "0.75");
        put("BD", "0.74");
        put("BA", "0.73");
        put("B8", "0.72");
        put("B5", "0.71");
        put("B3", "0.70");
        put("B0", "0.69");
        put("AD", "0.68");
        put("AB", "0.67");
        put("A8", "0.66");
        put("A6", "0.65");
        put("A3", "0.64");
        put("A1", "0.63");
        put("9E", "0.62");
        put("9C", "0.61");
        put("99", "0.60");
        put("96", "0.59");
        put("94", "0.58");
        put("91", "0.57");
        put("8F", "0.56");
        put("8C", "0.55");
        put("8A", "0.54");
        put("87", "0.53");
        put("85", "0.52");
        put("82", "0.51");
        put("80", "0.50");
        put("7D", "0.49");
        put("7A", "0.48");
        put("78", "0.47");
        put("75", "0.46");
        put("73", "0.45");
        put("70", "0.44");
        put("6E", "0.43");
        put("6B", "0.42");
        put("69", "0.41");
        put("66", "0.40");
        put("63", "0.39");
        put("61", "0.38");
        put("5E", "0.37");
        put("5C", "0.36");
        put("59", "0.35");
        put("57", "0.34");
        put("54", "0.33");
        put("52", "0.32");
        put("4F", "0.31");
        put("4D", "0.30");
        put("4A", "0.29");
        put("47", "0.28");
        put("45", "0.27");
        put("42", "0.26");
        put("40", "0.25");
        put("3D", "0.24");
        put("3B", "0.23");
        put("38", "0.22");
        put("36", "0.21");
        put("33", "0.20");
        put("30", "0.19");
        put("2E", "0.18");
        put("2B", "0.17");
        put("29", "0.16");
        put("26", "0.15");
        put("24", "0.14");
        put("21", "0.13");
        put("1F", "0.12");
        put("1C", "0.11");
        put("1A", "0.10");
        put("17", "0.09");
        put("14", "0.08");
        put("12", "0.07");
        put("0F", "0.06");
        put("0D", "0.05");
        put("0A", "0.04");
        put("08", "0.03");
        put("05", "0.02");
        put("03", "0.01");
        put("00", "0.00");
    }};

    /**
     * hex转透明度
     *
     * @param hex hex
     */
    public static String hexToAlpha(String hex) {
        return hexAlphaMap.getOrDefault(hex, "1");
    }

    /**
     * argbHex转rgba
     *
     * @param hex hex
     */
    public static String argbHexToRgba(String hex) {
        String aHex = hex.substring(0, 2);
        int r = Integer.parseInt(hex.substring(2, 4), 16);
        int g = Integer.parseInt(hex.substring(4, 6), 16);
        int b = Integer.parseInt(hex.substring(6, 8), 16);
        return "rgb(" + r + ", " + g + ", " + b + ", " + hexToAlpha(aHex) + ")";
    }
}
