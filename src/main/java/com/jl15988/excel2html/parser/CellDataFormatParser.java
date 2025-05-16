package com.jl15988.excel2html.parser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.HashMap;
import java.util.Map;

/**
 * 单元格数据格式解析器
 *
 * @author Jalon
 * @since 2025/5/16 9:14
 **/
public class CellDataFormatParser {

    /**
     * 单元格数据格式映射（仅适用于国内）
     * <p>
     * 来源于最新版 WPS 自定义格式
     */
    private static final Map<Short, String> CellDataFormatMap = new HashMap<Short, String>() {{
        // todo 部分不固定，5-8，23-26，41-44，176-...
        // 通用/数字格式
        // 常规
        put((short) 0, "General");
        // 整数
        put((short) 1, "0");
        // 两位小数
        put((short) 2, "0.00");
        // 千位分隔的整数
        put((short) 3, "#,##0");
        // 千位分隔的两位小数
        put((short) 4, "#,##0.00");
        // 带人民币符号的正负整数
        put((short) 5, "￥#,##0;￥-#,##0");
        // 带人民币符号的正负整数（负数为红色）
        put((short) 6, "￥#,##0;[红色]￥-#,##0");
        // 带人民币符号的正负两位小数
        put((short) 7, "￥#,##0.00;￥-#,##0.00");
        // 带人民币符号的正负两位小数（负数为红色）
        put((short) 8, "￥#,##0.00;[红色]￥-#,##0.00");
        // 百分比形式（整数）
        put((short) 9, "0%");
        // 百分比形式（两位小数）
        put((short) 10, "0.00%");
        // 科学计数法
        put((short) 11, "0.00E+00");

        // 日期格式
        // 年/月/日
        put((short) 14, "yyyy/m/d");
        // 日-月简写-年简写
        put((short) 15, "d-mmm-yy");
        // 日-月简写
        put((short) 16, "d-mmm");
        // 月简写-年简写
        put((short) 17, "mmm-yy");

        // 时间格式
        // 时:分
        put((short) 20, "h:mm");
        // 时:分:秒
        put((short) 21, "h:mm:ss");
        // 年/月/日 时:分
        put((short) 22, "yyyy/m/d h:mm");

        // 货币格式
        // 美元格式（整数）
        put((short) 23, "$#,##0_);($#,##0)");
        // 美元格式（整数，负数为红色）
        put((short) 24, "$#,##0_);[红色]($#,##0)");
        // 美元格式（两位小数）
        put((short) 25, "$#,##0.00_);($#,##0.00)");
        // 美元格式（两位小数，负数为红色）
        put((short) 26, "$#,##0.00_);[红色]($#,##0.00)");

        // 日期格式
        // 月/日/年简写
        put((short) 30, "m/d/yy");
        // 中文年月日
        put((short) 31, "yyyy\"年\"m\"月\"d\"日\"");

        // 时间格式
        // 中文时分
        put((short) 32, "h\"时\"mm\"分\"");
        // 中文时分秒
        put((short) 33, "h\"时\"mm\"分\"ss\"秒\"");

        // 数字格式
        // 正负整数
        put((short) 37, "#,##0;-#,##0");
        // 正负整数（负数为红色）
        put((short) 38, "#,##0;[红色]-#,##0");
        // 正负两位小数
        put((short) 39, "#,##0.00;-#,##0.00");
        // 正负两位小数（负数为红色）
        put((short) 40, "#,##0.00;[红色]-#,##0.00");

        // 会计专用格式
        // 会计格式（整数）
        put((short) 41, "_ * #,##0_ ;_ * -#,##0_ ;_ * \"-\"_ ;_ @_ ");
        // 会计格式（带人民币符号的整数）
        put((short) 42, "_ ￥* #,##0_ ;_ ￥* -#,##0_ ;_ ￥* \"-\"_ ;_ @_ ");
        // 会计格式（两位小数）
        put((short) 43, "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @_ ");
        // 会计格式（带人民币符号的两位小数）
        put((short) 44, "_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* \"-\"??_ ;_ @_ ");

        // 时间格式
        // 分:秒
        put((short) 45, "mm:ss");
        // [小时]:分:秒
        put((short) 46, "[h]:mm:ss");
        // 分:秒.0
        put((short) 47, "mm:ss.0");

        // 特殊格式
        // 科学计数法（简洁版）
        put((short) 48, "##0.0E+0");
        // 文本格式
        put((short) 49, "@");

        // 中文时间格式
        // 上午/下午时分
        put((short) 55, "上午/下午h\"时\"mm\"分\"");
        // 上午/下午时分秒
        put((short) 56, "上午/下午h\"时\"mm\"分\"ss\"秒\"");

        // 中文日期格式
        // 年月
        put((short) 57, "yyyy\"年\"m\"月\"");
        // 月日
        put((short) 58, "m\"月\"d\"日\"");

        // -------- 后面的可能不固定 --------

//        // 分数格式
//        // 一位分数
//        put((short) 176, "# ?/?");
//        // 两位分数
//        put((short) 177, "# ??/??");
//
//        // 中文日期格式（带数字样式）
//        // 中文数字年月日
//        put((short) 178, "[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\"");
//        // 中文数字年月
//        put((short) 179, "[DBNum1][$-804]yyyy\"年\"m\"月\"");
//        // 中文数字月日
//        put((short) 180, "[DBNum1][$-804]m\"月\"d\"日\"");
//
//        // 星期格式
//        // 星期全称
//        put((short) 181, "[$-804]aaaa");
//        // 星期简称
//        put((short) 182, "[$-804]aaa");
//
//        // 日期时间组合格式
//        // 年/月/日 时:分 AM/PM
//        put((short) 183, "yyyy/m/d h:mm AM/PM");
//        // 年简写/月/日
//        put((short) 184, "yy/m/d");
//        // 月/日
//        put((short) 185, "m/d");
//        // 月月/日日/年年
//        put((short) 186, "mm/dd/yy");
//        // 日日-月月简写-年年
//        put((short) 187, "dd-mmm-yy");
//        // 月份全称-年简写
//        put((short) 188, "mmmm-yy");
//        // 月份最简写
//        put((short) 189, "mmmmm");
//        // 月份最简写-年简写
//        put((short) 190, "mmmmm-yy");
//
//        // 12小时制时间格式
//        // 时:分 AM/PM
//        put((short) 191, "h:mm AM/PM");
//        // 时:分:秒 AM/PM
//        put((short) 192, "h:mm:ss AM/PM");
//
//        // 中文数字时间格式
//        // 中文数字时分
//        put((short) 193, "[DBNum1]h\"时\"mm\"分\"");
//        // 中文数字上午/下午时分
//        put((short) 194, "[DBNum1]上午/下午h\"时\"mm\"分\"");
//
//        // 人民币货币格式（使用¥符号）
//        // 带¥符号的正负整数
//        put((short) 195, "¥#,##0;¥-#,##0");
//        // 带¥符号的正负整数（负数为红色）
//        put((short) 196, "¥#,##0;[红色]¥-#,##0");
//        // 带¥符号的正负两位小数
//        put((short) 197, "¥#,##0.00;¥-#,##0.00");
//        // 带¥符号的正负两位小数（负数为红色）
//        put((short) 198, "¥#,##0.00;[红色]¥-#,##0.00");
    }};

    /**
     * 获取单元格数据格式
     *
     * @param cell 单元格
     * @return 单元格数据格式
     */
    public static String getDataFormatString(Cell cell) {
        // 获取数据格式编号
        CellStyle cellStyle = cell.getCellStyle();
        short dataFormat = cellStyle.getDataFormat();

        // 非 XSSF 直接默认返回
        if (!(cell instanceof XSSFCell)) {
            return cellStyle.getDataFormatString();
        }
        // 从工作簿获取样式资源
        XSSFRow row = (XSSFRow) cell.getRow();
        StylesTable stylesSource = row.getSheet().getWorkbook().getStylesSource();

        // 先从样式资源获取
        String fmt = stylesSource.getNumberFormatAt(dataFormat);
        if (fmt == null) {
            // 为空时从自定义映射获取
            if (dataFormat < 0 || dataFormat >= CellDataFormatMap.size()) {
                return null;
            }
            return CellDataFormatMap.get(dataFormat);
        }
        return fmt;
    }
}
