package com.jl15988.excel2html;

import com.jl15988.excel2html.converter.UnitConverter;
import com.jl15988.excel2html.parser.CellEmbedFileParser;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
        double totalMargin = UnitConverter.convert().convertInchToPoints(topMargin + bottomMargin);
        double A4_height = UnitConverter.convert().convertCMToPoints(29.7);
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
}
