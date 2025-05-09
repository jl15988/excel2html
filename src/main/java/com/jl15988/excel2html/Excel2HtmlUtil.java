package com.jl15988.excel2html;

import com.jl15988.excel2html.constant.UnitConstant;
import com.jl15988.excel2html.converter.FontSizeConverter;
import com.jl15988.excel2html.model.unit.UnitInch;
import com.jl15988.excel2html.model.unit.UnitMillimetre;
import com.jl15988.excel2html.model.unit.UnitPoint;
import com.jl15988.excel2html.parser.CellEmbedFileParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * Excel 转 HTML 工具类
 * <p>
 * 提供一系列工具方法，用于辅助Excel到HTML的转换过程。
 * 包含Excel文件解析、单位换算、尺寸计算等功能。
 * </p>
 *
 * @author Jalon
 * @since 2024/12/6 10:01
 **/
public class Excel2HtmlUtil {

    /**
     * 从Excel文件数据中加载嵌入文件（如图片）
     * <p>
     * Excel文件实际上是一个ZIP格式的压缩包，此方法解析压缩包内容，
     * 提取嵌入的图片等多媒体文件。
     * </p>
     *
     * @param fileData Excel文件的字节数据
     * @return 嵌入文件映射，key为文件ID，value为图片数据
     * @throws IOException 如果文件解析过程中出错
     */
    public static Map<String, XSSFPictureData> doLoadEmbedFile(byte[] fileData) throws IOException {
        if (Objects.nonNull(fileData)) {
            // 解析嵌入图片数据
            Map<String, String> stringStringMap = CellEmbedFileParser.processZipEntries(new ByteArrayInputStream(fileData));
            return CellEmbedFileParser.processPictures(new ByteArrayInputStream(fileData), stringStringMap);
        }
        return new HashMap<>();
    }

    /**
     * 获取工作表中的最大行数
     * <p>
     * 返回工作表中最后一行的索引值+1，即实际行数
     * </p>
     *
     * @param sheet 工作表对象
     * @return 工作表中的最大行数
     */
    public static int getMaxRowNum(Sheet sheet) {
        return sheet.getLastRowNum() + 1;
    }

    /**
     * 获取工作表中的最大列数
     * <p>
     * 通过遍历所有行，找出最大的列索引值
     * </p>
     *
     * @param sheet 工作表对象
     * @return 工作表中的最大列数
     */
    public static int getMaxColNum(Sheet sheet) {
        short colNum = 0;
        for (Row row : sheet) {
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum > colNum) {
                colNum = lastCellNum;
            }
        }
        return colNum;
    }

    /**
     * 获取打印页的最后一行索引
     * <p>
     * 根据纸张高度计算可显示的最大行数。通过计算行的累计高度，
     * 确定在指定纸张高度下最多可显示到哪一行。
     * 注意：通过计算行、列占用截取行、列，计算结果可能不太准确。
     * </p>
     *
     * @param sheet       工作表对象
     * @param paperHeight 纸张高度，单位毫米
     * @return 打印页的最后一行索引
     */
    public static int getPrintLastRowNum(Sheet sheet, Float paperHeight) {
        XSSFPrintSetup printSetup = (XSSFPrintSetup) sheet.getPrintSetup();

        // 获取页边距
        double topMargin = printSetup.getTopMargin();
        double bottomMargin = printSetup.getBottomMargin();
        double totalVerticalMargin = new UnitInch(topMargin + bottomMargin).toPoint().getValue();

        // 转换为点
        double paperHeightPoints = new UnitMillimetre(paperHeight).toPoint().getValue();

        // 考虑一些调整空间，避免精确的边缘情况
        double thresholdValue = 0;
        double overHeight = paperHeightPoints - totalVerticalMargin - thresholdValue;

        int lastRowNum = sheet.getLastRowNum();
        double defaultRowHeightInPoints = (double) sheet.getDefaultRowHeightInPoints();
        int currentRowNum = 0;

        double totalHeight = 0;
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            double rowHeight = Objects.nonNull(row) ? (double) row.getHeightInPoints() : defaultRowHeightInPoints;
            totalHeight += rowHeight;
            if (totalHeight > overHeight) {
                // 判断差值是否超过最后一行一半高度
                double difference = totalHeight - overHeight;
                System.out.println(rowHeight);
                System.out.println(difference);
                if (difference > rowHeight / 2) {
                    currentRowNum = rowIndex + 1;
                }
                break;
            } else {
                currentRowNum = rowIndex + 1;
            }
        }

        return currentRowNum;
    }

    /**
     * 获取打印页的最后一列索引
     * <p>
     * 根据纸张宽度计算可显示的最大列数。通过计算列的累计宽度，
     * 确定在指定纸张宽度下最多可显示到哪一列。
     * 注意：通过计算行、列占用截取行、列，计算结果可能不太准确。
     * </p>
     *
     * @param sheet      工作表对象
     * @param paperWidth 纸张宽度，单位毫米
     * @return 打印页的最后一列索引
     */
    public static int getPrintLastColNum(Sheet sheet, Float paperWidth) {
        XSSFPrintSetup printSetup = (XSSFPrintSetup) sheet.getPrintSetup();

        // 获取页边距
        double leftMargin = printSetup.getLeftMargin();
        double rightMargin = printSetup.getRightMargin();
        System.out.println(leftMargin + " " + rightMargin);
        double totalHorizontalMargin = new UnitInch(leftMargin + rightMargin).toPoint().getValue();

        // 转换为点
        double paperWidthPoints = new UnitMillimetre(paperWidth).toPoint().getValue();

        // 考虑一些调整空间，避免精确的边缘情况
        double thresholdValue = 0;
        double overWidth = paperWidthPoints - totalHorizontalMargin - thresholdValue;

        // 获取最大列数
        int maxColNum = getMaxColNum(sheet);
        int currentColNum = 0;

        double totalWidth = 0;
        for (int colIndex = 0; colIndex < maxColNum; colIndex++) {
            // 获取列宽（以磅为单位）
            int colWidthInPixels = getColumnWidthInPixels(sheet, colIndex);
            // 将像素转换为磅
            double colWidthInPoints = new UnitPoint(colWidthInPixels, UnitConstant.DEFAULT_DPI).getValue();

            totalWidth += colWidthInPoints;
            if (totalWidth > overWidth) {
                // 判断差值是否超过最后一列一半宽度
                double difference = totalWidth - overWidth;
                if (difference > colWidthInPoints / 2) {
                    currentColNum = colIndex;
                }
                break;
            } else {
                currentColNum = colIndex + 1;
            }
        }

        return currentColNum;
    }

    /**
     * 获取工作簿默认字体的像素大小
     * <p>
     * 根据工作簿默认字体的名称和高度计算对应的像素大小
     * </p>
     *
     * @param workbook 工作簿对象
     * @return 默认字体的像素大小
     */
    public static double getDefaultFontPixelSize(Workbook workbook) {
        Font defaultWorkbookFont = Excel2HtmlUtil.getDefaultWorkbookFont(workbook);
        return FontSizeConverter.getPixelSize(defaultWorkbookFont.getFontName(), defaultWorkbookFont.getFontHeightInPoints());
    }

    /**
     * 获取默认列宽（像素值）
     * <p>
     * 计算工作簿默认列宽的像素值，考虑了字体大小、内边距和网格线宽度
     * </p>
     *
     * @param workbook 工作簿对象
     * @return 默认列宽的像素值
     */
    public static int getDefaultColumnWidthInPixels(Workbook workbook) {
        double defaultColumnCharWidth = 8;
        double dfw = getDefaultFontPixelSize(workbook);
        // 两边的内边距
        int padding = (int) Math.ceil((double) dfw / 4);
        // 网格线宽度
        double lineWidth = 1.0;
        // 计算列宽
        double columnWidth = defaultColumnCharWidth * dfw + padding * 2 + lineWidth;
        // 最终向上取8的整数倍数（Excel规定每列的像素宽度必须是8的整数倍（为了在滚动时提高渲染性能））
        return (int) Math.ceil(columnWidth / 8.0) * 8;
    }

    /**
     * 获取默认字符列宽
     * <p>
     * 计算工作簿默认列宽的字符数
     * </p>
     *
     * @param workbook 工作簿对象
     * @return 默认列宽的字符数
     */
    public static int getDefaultColumnWidth(Workbook workbook) {
        double defaultFontPixelSize = Excel2HtmlUtil.getDefaultFontPixelSize(workbook);
        return (int) Math.ceil(getDefaultColumnWidthInPixels(workbook) / defaultFontPixelSize);
    }

    /**
     * 获取特殊格式的默认字符列宽
     * <p>
     * 返回负数，且扩大10000倍的值，用于标识特殊列宽并保留小数精度
     * </p>
     *
     * @param workbook 工作簿对象
     * @return 特殊格式的默认字符列宽
     */
    public static int getDefaultColumnWidthSpecial(Workbook workbook) {
        double defaultFontPixelSize = Excel2HtmlUtil.getDefaultFontPixelSize(workbook);
        return -(int) Math.ceil(getDefaultColumnWidthInPixels(workbook) / defaultFontPixelSize * 10000);
    }

    /**
     * 获取指定列的像素宽度
     * <p>
     * 根据工作表中列宽的设置，计算对应的像素宽度
     * </p>
     *
     * @param sheet       工作表对象
     * @param columnIndex 列索引
     * @return 列的像素宽度
     */
    public static int getColumnWidthInPixels(Sheet sheet, int columnIndex) {
        double defaultFontPixelSize = Excel2HtmlUtil.getDefaultFontPixelSize(sheet.getWorkbook());
        int columnWidth = sheet.getColumnWidth(columnIndex);
        if (columnWidth < 0) {
            return (int) Math.ceil(((double) (-columnWidth / 10000) / 256 * defaultFontPixelSize));
        }
        return (int) ((double) columnWidth / 256 * defaultFontPixelSize);
    }

    /**
     * 获取工作簿默认的字体
     * <p>
     * 如果工作簿中已定义字体，则返回第一个字体；
     * 否则创建一个新的默认字体
     * </p>
     *
     * @param workbook 工作簿对象
     * @return 工作簿默认字体
     */
    public static Font getDefaultWorkbookFont(Workbook workbook) {
        int numberOfFonts = workbook.getNumberOfFonts();
        if (numberOfFonts > 0) {
            return workbook.getFontAt(0);
        }
        Font newFont = workbook.createFont();
        newFont.setFontName(Excel2Html.DEFAULT_ALTERNATE_FONT_FAMILY);
        newFont.setFontHeightInPoints((short) 11);
        return newFont;
    }

    /**
     * 获取增强的数据格式字符串
     * 补充POI中BuiltinFormats类中缺失的日期格式
     *
     * @param cell 单元格
     * @return 格式字符串
     */
    public static String getDataFormatString(Cell cell) {
        if (cell == null) {
            return null;
        }

        CellStyle style = cell.getCellStyle();
        if (style == null) {
            return null;
        }

        short dataFormat = style.getDataFormat();

        // 补充常见的内置日期格式
        Map<Short, String> additionalFormats = new HashMap<>();

        // 通用/数字格式
        additionalFormats.put((short) 0, "General");                   // 0: 常规
        additionalFormats.put((short) 1, "0");                         // 1: 整数
        additionalFormats.put((short) 2, "0.00");                      // 2: 两位小数
        additionalFormats.put((short) 3, "#,##0");                     // 3: 千分位整数
        additionalFormats.put((short) 4, "#,##0.00");                  // 4: 千分位两位小数
        additionalFormats.put((short) 5, "$#,##0_);($#,##0)");         // 5: 会计格式（美元符号）
        additionalFormats.put((short) 6, "$#,##0_);[Red]($#,##0)");    // 6: 会计格式（美元符号，负数红色）
        additionalFormats.put((short) 7, "$#,##0.00_);($#,##0.00)");   // 7: 会计格式（美元符号，两位小数）
        additionalFormats.put((short) 8, "$#,##0.00_);[Red]($#,##0.00)"); // 8: 会计格式（美元符号，两位小数，负数红色）
        additionalFormats.put((short) 9, "0%");                        // 9: 百分比，整数
        additionalFormats.put((short) 10, "0.00%");                    // 10: 百分比，两位小数
        additionalFormats.put((short) 11, "0.00E+00");                 // 11: 科学计数法
        additionalFormats.put((short) 12, "# ?/?");                    // 12: 分数
        additionalFormats.put((short) 13, "# ??/??");                  // 13: 分数

        // 日期格式
        additionalFormats.put((short) 14, "m/d/yy");                   // 14: 短日期 mm-dd-yy
        additionalFormats.put((short) 15, "d-mmm-yy");                 // 15: 长日期 dd-mmm-yy
        additionalFormats.put((short) 16, "d-mmm");                    // 16: dd-mmm
        additionalFormats.put((short) 17, "mmm-yy");                   // 17: mmm-yy
        additionalFormats.put((short) 18, "h:mm AM/PM");               // 18: 时间 h:mm AM/PM
        additionalFormats.put((short) 19, "h:mm:ss AM/PM");            // 19: 时间 h:mm:ss AM/PM
        additionalFormats.put((short) 20, "h:mm");                     // 20: 时间 h:mm
        additionalFormats.put((short) 21, "h:mm:ss");                  // 21: 时间 h:mm:ss
        additionalFormats.put((short) 22, "m/d/yy h:mm");              // 22: 日期时间 m/d/yy h:mm

        // Office Excel中文版常用日期格式
        additionalFormats.put((short) 27, "[$-404]e/m/d");             // 27: 农历日期格式
        additionalFormats.put((short) 28, "[$-404]e\"年\"m\"月\"d\"日\""); // 28: 中文农历日期
        additionalFormats.put((short) 29, "[$-404]e\"年\"m\"月\"");     // 29: 中文农历年月
        additionalFormats.put((short) 30, "m-d-yy");                   // 30: 日期格式 m-d-yy
        additionalFormats.put((short) 31, "yyyy\"年\"m\"月\"d\"日\"");   // 31: 中文日期年月日
        additionalFormats.put((short) 32, "h\"时\"mm\"分\"");           // 32: 中文时间时分
        additionalFormats.put((short) 33, "h\"时\"mm\"分\"ss\"秒\"");    // 33: 中文时间时分秒
        additionalFormats.put((short) 34, "上午/下午h\"时\"mm\"分\"");    // 34: 中文带上下午时间
        additionalFormats.put((short) 35, "上午/下午h\"时\"mm\"分\"ss\"秒\""); // 35: 中文带上下午时间秒
        additionalFormats.put((short) 36, "yyyy\"年\"m\"月\"");         // 36: 中文年月
        additionalFormats.put((short) 37, "m\"月\"d\"日\"");           // 37: 中文月日
        additionalFormats.put((short) 38, "yyyy-mm-dd");              // 38: ISO标准日期
        additionalFormats.put((short) 39, "yyyy\"年\"m\"月\"d\"日\" h\"时\"mm\"分\"ss\"秒\""); // 39: 中文完整日期时间
        additionalFormats.put((short) 40, "yyyy/m/d h:mm");           // 40: 日期时间 yyyy/m/d h:mm
        additionalFormats.put((short) 41, "yyyy/m/d h:mm:ss");        // 41: 日期时间带秒
        additionalFormats.put((short) 42, "yyyy-mm-dd hh:mm:ss");     // 42: ISO标准日期时间

        // 货币格式
        additionalFormats.put((short) 43, "\"￥\"#,##0;\"￥\"\\-#,##0");  // 43: 人民币整数
        additionalFormats.put((short) 44, "\"￥\"#,##0;[Red]\"￥\"\\-#,##0"); // 44: 人民币整数(负数红色)
        additionalFormats.put((short) 45, "\"￥\"#,##0.00;\"￥\"\\-#,##0.00"); // 45: 人民币两位小数
        additionalFormats.put((short) 46, "\"￥\"#,##0.00;[Red]\"￥\"\\-#,##0.00"); // 46: 人民币两位小数(负数红色)
        additionalFormats.put((short) 47, "\"$\"#,##0.00_);\"$\"#,##0.00\\)"); // 47: 美元两位小数(负括号)
        additionalFormats.put((short) 48, "\"$\"#,##0.00_);[Red]\"$\"#,##0.00\\)"); // 48: 美元两位小数(负红色括号)

        // 会计专用
        additionalFormats.put((short) 49, "_ * #,##0_ ;_ * \\-#,##0_ ;_ * \"-\"_ ;_ @_ "); // 49: 会计整数
        additionalFormats.put((short) 50, "_ \"￥\"* #,##0_ ;_ \"￥\"* \\-#,##0_ ;_ \"￥\"* \"-\"_ ;_ @_ "); // 50: 会计人民币整数
        additionalFormats.put((short) 51, "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * \"-\"??_ ;_ @_ "); // 51: 会计两位小数
        additionalFormats.put((short) 52, "_ \"￥\"* #,##0.00_ ;_ \"￥\"* \\-#,##0.00_ ;_ \"￥\"* \"-\"??_ ;_ @_ "); // 52: 会计人民币两位小数

        // 自定义格式
        additionalFormats.put((short) 57, "yyyy/mm/dd;@");            // 57: 日期(无格式文本)
        additionalFormats.put((short) 58, "yyyy/m/d;@");              // 58: 日期简化(无格式文本)
        additionalFormats.put((short) 59, "dd/mm/yyyy;@");            // 59: 日期欧洲格式(无格式文本)
        additionalFormats.put((short) 67, "yyyy\"年\"m\"月\"d\"日\";@");// 67: 中文日期(无格式文本)

        // 从自定义的格式映射中获取
        String enhancedFormat = additionalFormats.get(dataFormat);
        if (enhancedFormat != null) {
            return enhancedFormat;
        }

        // 如果自定义映射中没有，再获取原始的格式字符串
        String dataFormatString = style.getDataFormatString();
        if (dataFormatString != null && !dataFormatString.isEmpty() && !dataFormatString.equals("General")) {
            return dataFormatString;
        }

        // 如果都没有匹配到，返回通用格式
        return "General";
    }

    /**
     * 增强版日期格式判断方法，解决POI的DateUtil.isCellDateFormatted方法不准确的问题
     *
     * @param cell 单元格
     * @return 是否为日期格式
     */
    public static boolean isCellDateFormatted(Cell cell) {
        if (cell == null) {
            return false;
        }

        // 先使用原始方法判断
        if (DateUtil.isCellDateFormatted(cell)) {
            return true;
        }

        // 如果原始判断为否，进一步判断
        CellStyle style = cell.getCellStyle();
        if (style == null) {
            return false;
        }

        short dataFormat = style.getDataFormat();
        String dataFormatString = style.getDataFormatString();

        // 已知的日期格式代码列表
        short[] dateCodes = new short[]{
                14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31,
                32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 57, 58, 59, 67
        };

        // 检查格式代码是否在已知日期格式列表中
        for (short dateCode : dateCodes) {
            if (dataFormat == dateCode) {
                return true;
            }
        }

        // 检查格式字符串是否含有日期格式的特征
        if (dataFormatString != null) {
            String lower = dataFormatString.toLowerCase();

            // 检查是否包含常见日期格式关键字
            if (lower.contains("y") || lower.contains("m") || lower.contains("d") ||
                    lower.contains("h") || lower.contains("s") || lower.contains("年") ||
                    lower.contains("月") || lower.contains("日") || lower.contains("时") ||
                    lower.contains("分") || lower.contains("秒") || lower.contains("am/pm") ||
                    lower.contains("a/p")) {

                // 额外检查，排除可能是时间但不是日期的格式
                boolean hasDatePart = lower.contains("y") || lower.contains("m") ||
                        lower.contains("d") || lower.contains("年") ||
                        lower.contains("月") || lower.contains("日");

                // 排除科学计数法格式，它们可能包含"e"，容易被错误识别为日期
                boolean notScientific = !lower.contains("e+") && !lower.contains("e-");

                // 排除包含货币符号的格式
                boolean notCurrency = !lower.contains("$") && !lower.contains("￥") &&
                        !lower.contains("rmb") && !lower.contains("cny");

                return hasDatePart && notScientific && notCurrency;
            }
        }

        // 针对单元格类型为数值且值在合理日期范围内的特殊处理
        if (cell.getCellType() == CellType.NUMERIC) {
            double value = cell.getNumericCellValue();
            // Excel日期从1900-01-01开始计算，值为1
            // 检查值是否在合理日期范围内（1900-01-01到2099-12-31之间）
            if (value >= 1 && value <= 73050) { // 约等于2099-12-31
                // 获取增强的数据格式
                String enhancedFormat = getDataFormatString(cell);
                if (enhancedFormat != null && !enhancedFormat.equals("General") && !enhancedFormat.contains("#")) {
                    // 如果不是General且不含有数字格式标记#，可能是日期
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * 将Excel日期格式字符串转换为Java SimpleDateFormat格式字符串
     *
     * @param excelFormat Excel日期格式字符串
     * @return Java SimpleDateFormat格式字符串
     */
    public static String convertExcelDateFormatToJava(String excelFormat) {
        if (excelFormat == null || excelFormat.isEmpty()) {
            return "yyyy-MM-dd";
        }

        // 开始进行通用转换
        // 转换年份格式
        String javaFormat = excelFormat;

        // 处理可能包含的语言区域代码，如[$-804]等
        javaFormat = javaFormat.replaceAll("\\[\\$-[0-9A-F]+\\]", "");

        // 替换单引号，在Java SimpleDateFormat中单引号是特殊字符
        javaFormat = javaFormat.replace("'", "''");

        // 替换AM/PM
        javaFormat = javaFormat.replace("AM/PM", "a");
        javaFormat = javaFormat.replace("上午/下午", "a");

        // 替换年份格式
        javaFormat = javaFormat.replace("yyyy", "yyyy")
                .replace("yy", "yy");

        // 替换月份格式，必须在替换日之前处理，避免m被误认为是分钟
        javaFormat = javaFormat.replace("mmm", "MMM")
                .replace("mm", "MM")
                .replace("m", "M");

        // 替换日格式
        javaFormat = javaFormat.replace("dddd", "EEEE")
                .replace("ddd", "EEE")
                .replace("dd", "dd")
                .replace("d", "d");

        // 替换小时格式
        javaFormat = javaFormat.replace("hh", "HH")
                .replace("h", "H");

        // 替换星期格式
        javaFormat = javaFormat.replace("aaa", "EEE")
                .replace("aaaa", "EEEE");

        // 替换常见的双引号包围的文本为单引号包围
        javaFormat = javaFormat.replace("\"年\"", "'年'")
                .replace("\"月\"", "'月'")
                .replace("\"日\"", "'日'")
                .replace("\"时\"", "'时'")
                .replace("\"分\"", "'分'")
                .replace("\"秒\"", "'秒'")
                .replace("\"星期\"", "'星期'");

        return javaFormat;
    }
}
