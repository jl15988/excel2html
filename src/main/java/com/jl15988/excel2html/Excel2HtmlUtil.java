package com.jl15988.excel2html;

import com.jl15988.excel2html.converter.FontSizeConverter;
import com.jl15988.excel2html.model.unit.Inch;
import com.jl15988.excel2html.model.unit.Millimetre;
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
 * excel 转 html 工具
 *
 * @author Jalon
 * @since 2024/12/6 10:01
 **/
public class Excel2HtmlUtil {

    /**
     * 加载嵌入文件
     *
     * @throws IOException
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
     * 获取最大行数
     *
     * @param sheet sheet
     */
    public static int getMaxRowNum(Sheet sheet) {
        return sheet.getLastRowNum() + 1;
    }

    /**
     * 获取最大列数
     *
     * @param sheet sheet
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
     * 获取打印页的最后一行
     */
    public static int getPrintLastRowNum(Sheet sheet) {
        //        PrintSetup.A4_PAPERSIZE
//        String printArea = workbook.getPrintArea(0);
        XSSFPrintSetup printSetup = (XSSFPrintSetup) sheet.getPrintSetup();
        double topMargin = printSetup.getTopMargin();
        double bottomMargin = printSetup.getBottomMargin();
        double totalMargin = new Inch(topMargin + bottomMargin).toPoint().getValue();
        double A4_height = new Millimetre(297).toPoint().getValue();
//        double thresholdValue = 18;
//        double thresholdValue = 15;
        double thresholdValue = 0;
        double overHeight = A4_height - totalMargin - thresholdValue;

        int lastRowNum = sheet.getLastRowNum();
        double defaultRowHeightInPoints = (double) sheet.getDefaultRowHeightInPoints();
        int currentRowNum = 0;

        double totalHeight = 0;
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (Objects.nonNull(row)) {
                totalHeight += (double) row.getHeightInPoints();
            } else {
                totalHeight += defaultRowHeightInPoints;
            }
            if (totalHeight > overHeight) {
                break;
            }
            currentRowNum = rowIndex + 1;
        }

        return currentRowNum;
    }

    /**
     * 获取工作簿默认字体像素大小
     *
     * @param workbook 工作簿
     */
    public static double getDefaultFontPixelSize(Workbook workbook) {
        Font defaultWorkbookFont = Excel2HtmlUtil.getDefaultWorkbookFont(workbook);
        return FontSizeConverter.getPixelSize(defaultWorkbookFont.getFontName(), defaultWorkbookFont.getFontHeightInPoints());
    }

    /**
     * 获取默认像素列宽
     *
     * @param workbook 工作簿
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
     *
     * @param workbook 工作簿
     */
    public static int getDefaultColumnWidth(Workbook workbook) {
        double defaultFontPixelSize = Excel2HtmlUtil.getDefaultFontPixelSize(workbook);
        return (int) Math.ceil(getDefaultColumnWidthInPixels(workbook) / defaultFontPixelSize);
    }

    /**
     * 获取默认字符列宽，获取的为负数，且扩大 10000 倍的，用于标识和保留小数
     *
     * @param workbook 工作簿
     */
    public static int getDefaultColumnWidthSpecial(Workbook workbook) {
        double defaultFontPixelSize = Excel2HtmlUtil.getDefaultFontPixelSize(workbook);
        return -(int) Math.ceil(getDefaultColumnWidthInPixels(workbook) / defaultFontPixelSize * 10000);
    }

    /**
     * 获取像素列宽
     *
     * @param sheet       sheet
     * @param columnIndex 列
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
     *
     * @param workbook 工作簿
     */
    public static Font getDefaultWorkbookFont(Workbook workbook) {
        int numberOfFonts = workbook.getNumberOfFonts();
        if (numberOfFonts > 0) {
            return workbook.getFontAt(0);
        }
        Font newFont = workbook.createFont();
        newFont.setFontName("宋体");
        newFont.setFontHeightInPoints((short) 11);
        return newFont;
    }
}
