package com.jl15988.excel2html.parser;

import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;

/**
 * 单元格格式化器
 *
 * @author Jalon
 * @since 2024/11/29 19:23
 **/
class CellValueFormatter {

    /**
     * 获格式化特殊数字
     *
     * @param cell   单元格
     * @param number 数字
     * @return 值
     */
    public static String formatNumericValue(Cell cell, double number) {
        String dataFormatString = cell.getCellStyle().getDataFormatString();
        if ("General".equals(dataFormatString)) {
            BigDecimal bigDecimal = new BigDecimal(number);
            // 去除末尾所有0
            BigDecimal strippedNumber = bigDecimal.stripTrailingZeros();
            // 如果小数位数为0，则取整数
            if (strippedNumber.scale() == 0) {
                return strippedNumber.toBigInteger().toString();
            }
        }
        return String.valueOf(number);
    }
}
