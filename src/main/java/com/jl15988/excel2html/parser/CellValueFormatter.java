package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.Excel2HtmlUtil;
import com.jl15988.excel2html.evaluators.CustomConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.math.BigDecimal;
import java.util.List;

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

    /**
     * 格式化日期内容，注意必须是数字日期格式才可使用
     *
     * @param cell       单元格
     * @param dateFormat 日期格式
     * @return 格式化后的内容
     */
    public static String formatDateValue(Cell cell, String dateFormat) {
//        try {
//            // 将Excel日期格式转换为Java SimpleDateFormat格式
//            String javaDateFormat = Excel2HtmlUtil.convertExcelDateFormatToJava(dateFormat);
//            // 使用SimpleDateFormat格式化日期
//            java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat(javaDateFormat);
//            return sdf.format(cell.getDateCellValue());
//        } catch (Exception e) {
//            // 如果格式化失败，回退到DataFormatter
//            DataFormatter dataFormatter = new DataFormatter();
//            dataFormatter.setUseCachedValuesForFormulaCells(true);
//            return dataFormatter.formatCellValue(cell, null, new CustomConditionalFormattingEvaluator(null, null));
//        }
        try {
            DataFormatter dataFormatter = new DataFormatter();
            dataFormatter.setUseCachedValuesForFormulaCells(true);
            return dataFormatter.formatCellValue(cell, null, new CustomConditionalFormattingEvaluator(cell.getRow().getSheet().getWorkbook(), null));
        } catch (Exception e) {
            return "";
        }
    }
}
