package com.jl15988.excel2html;

import com.jl15988.excel2html.html.HtmlPage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import java.io.IOException;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/1 14:24
 **/
public class Excel2HtmlUtil {

    /**
     * 表格转 html
     *
     * @param sheet         表格 sheet
     * @param columnNum     列数
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到 css
     */
    public static HtmlPage toHtml(Sheet sheet, int columnNum, boolean compressStyle) throws IOException {
        return new Excel2Html().setCompressStyle(compressStyle).buildHtml(sheet, columnNum);
    }

    /**
     * 表格转 html
     *
     * @param sheet         表格 sheet
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到 css
     */
    public static HtmlPage toHtml(Sheet sheet, boolean compressStyle) throws IOException {
        return new Excel2Html().setCompressStyle(compressStyle).buildHtml(sheet);
    }

    /**
     * 表格转 html
     *
     * @param sheet 表格 sheet
     */
    public static HtmlPage toHtml(Sheet sheet) throws IOException {
        return new Excel2Html().buildHtml(sheet);
    }

    /**
     * 表格转 html
     *
     * @param sheet         表格 sheet
     * @param columnNum     列数
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到 css
     * @param embedFileMap  嵌入文件映射
     */
    public static HtmlPage toHtml(Sheet sheet, int columnNum, boolean compressStyle, Map<String, XSSFPictureData> embedFileMap) throws IOException {
        return new Excel2Html().setLoadEmbedFile(compressStyle).setEmbedFileMap(embedFileMap).buildHtml(sheet, columnNum);
    }
}
