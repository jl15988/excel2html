package com.jl15988.excel2html;

import com.jl15988.excel2html.formatter.ICellValueFormater;
import com.jl15988.excel2html.html.HtmlPage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import java.io.IOException;
import java.util.Map;

/**
 * Excel 转 HTML 工具助手类
 * <p>
 * 提供一系列静态方法，简化Excel转HTML的调用过程。
 * 本类封装了对{@link Excel2Html}的调用，提供更简便的API接口。
 * </p>
 *
 * @author Jalon
 * @since 2024/12/1 14:24
 **/
public class Excel2HtmlHelper {

    /**
     * 将Excel表格转换为HTML
     * <p>
     * 可指定转换的行列范围和是否压缩样式
     * </p>
     *
     * @param sheet         Excel工作表对象
     * @param startRowIndex 开始行索引，可为null默认为0
     * @param endRowIndex   结束行索引，可为null默认为最后一行
     * @param startColIndex 开始列索引，可为null默认为0
     * @param endColIndex   结束列索引，可为null默认为最后一列
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到CSS类中
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex, boolean compressStyle) throws IOException {
        return new Excel2Html().setCompressStyle(compressStyle).buildHtml(sheet, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }

    /**
     * 将Excel表格转换为HTML
     * <p>
     * 可指定转换的行列范围，默认启用样式压缩
     * </p>
     *
     * @param sheet         Excel工作表对象
     * @param startRowIndex 开始行索引，可为null默认为0
     * @param endRowIndex   结束行索引，可为null默认为最后一行
     * @param startColIndex 开始列索引，可为null默认为0
     * @param endColIndex   结束列索引，可为null默认为最后一列
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex) throws IOException {
        return new Excel2Html().setCompressStyle(true).buildHtml(sheet, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }

    /**
     * 将整个Excel工作表转换为HTML
     * <p>
     * 可指定是否压缩样式
     * </p>
     *
     * @param sheet         Excel工作表对象
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到CSS类中
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet, boolean compressStyle) throws IOException {
        return new Excel2Html().setCompressStyle(compressStyle).buildHtml(sheet);
    }

    /**
     * 将整个Excel工作表转换为HTML
     * <p>
     * 默认启用样式压缩
     * </p>
     *
     * @param sheet Excel工作表对象
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet) throws IOException {
        return new Excel2Html().buildHtml(sheet);
    }

    /**
     * 将Excel表格转换为HTML，可自定义嵌入文件映射
     * <p>
     * 可指定转换的行列范围、是否压缩样式，以及嵌入文件映射
     * </p>
     *
     * @param sheet         Excel工作表对象
     * @param startRowIndex 开始行索引，可为null默认为0
     * @param endRowIndex   结束行索引，可为null默认为最后一行
     * @param startColIndex 开始列索引，可为null默认为0
     * @param endColIndex   结束列索引，可为null默认为最后一列
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到CSS类中
     * @param embedFileMap  嵌入文件映射，用于提供Excel中嵌入的图片等文件
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex, boolean compressStyle, Map<String, XSSFPictureData> embedFileMap) throws IOException {
        return new Excel2Html().setLoadEmbedFile(compressStyle).setEmbedFileMap(embedFileMap).buildHtml(sheet, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }

    /**
     * 将Excel表格转换为HTML，支持自定义单元格值格式化和嵌入文件映射
     * <p>
     * 可指定转换的行列范围、是否压缩样式、单元格值格式化器和嵌入文件映射
     * </p>
     *
     * @param sheet             Excel工作表对象
     * @param startRowIndex     开始行索引，可为null默认为0
     * @param endRowIndex       结束行索引，可为null默认为最后一行
     * @param startColIndex     开始列索引，可为null默认为0
     * @param endColIndex       结束列索引，可为null默认为最后一列
     * @param compressStyle     是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到CSS类中
     * @param cellValueFormater 单元格值格式化器，用于自定义单元格值的显示格式
     * @param embedFileMap      嵌入文件映射，用于提供Excel中嵌入的图片等文件
     * @return HTML页面对象
     * @throws IOException 如果处理过程中发生IO异常
     */
    public static HtmlPage toHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex, boolean compressStyle, ICellValueFormater cellValueFormater, Map<String, XSSFPictureData> embedFileMap) throws IOException {
        return new Excel2Html().setLoadEmbedFile(compressStyle).setCellValueFormater(cellValueFormater).setEmbedFileMap(embedFileMap).buildHtml(sheet, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }
}
