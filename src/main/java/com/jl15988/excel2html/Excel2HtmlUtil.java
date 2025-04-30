package com.jl15988.excel2html;

import com.jl15988.excel2html.constant.UnitConstant;
import com.jl15988.excel2html.converter.FontSizeConverter;
import com.jl15988.excel2html.model.unit.UnitInch;
import com.jl15988.excel2html.model.unit.UnitMillimetre;
import com.jl15988.excel2html.model.unit.UnitPoint;
import com.jl15988.excel2html.parser.CellEmbedFileParser;
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
}
