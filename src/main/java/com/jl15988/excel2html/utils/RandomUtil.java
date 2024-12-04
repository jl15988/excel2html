package com.jl15988.excel2html.utils;

/**
 * @author Jalon
 * @since 2024/12/2 10:23
 **/
public class RandomUtil {

    /**
     * 生成随机hash名
     *
     * @param length 长度
     */
    public static String randomHash(int length) {
        String chars = "abcdefghijklmnopqrstuvwxyz0123456789";
        StringBuilder result = new StringBuilder();
        int charsLength = chars.length();
        for (int i = 0; i < length; i++) {
            result.append(chars.charAt((int) Math.floor(Math.random() * charsLength)));
        }
        return result.toString();
    }

    /**
     * 生成随机hash名
     *
     * @param length 长度
     */
    public static String randomHash2(int length) {
        String chars = "abcdefghijklmnopqrstuvwxyz";
        StringBuilder result = new StringBuilder();
        int charsLength = chars.length();
        for (int i = 0; i < length; i++) {
            result.append(chars.charAt((int) Math.floor(Math.random() * charsLength)));
        }
        return result.toString();
    }
}
