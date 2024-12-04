package com.jl15988.excel2html;

import com.jl15988.excel2html.html.HtmlPage;
import com.jl15988.excel2html.parser.CellEmbedFileParser;
import com.jl15988.excel2html.utils.FileUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @author Jalon
 * @since 2024/12/4 9:13
 **/
public class Excel2HtmlBuilder {

    /**
     * 文件数据
     */
    private byte[] fileData;

    /**
     * 文件类型
     */
    private String fileType;

    /**
     * sheet-构建后的 html
     */
    private Map<Integer, HtmlPage> sheetToHtmlMap;

    /**
     * 是否启用压缩样式
     */
    private boolean isCompressStyle = true;

    /**
     * 是否加载嵌入文件
     */
    private boolean isLoadEmbedFile = true;

    /**
     * 是否去除单元格两端空格
     */
    private boolean isTrimCellValue;

    /**
     * 嵌入文件缓存
     */
    private Map<String, XSSFPictureData> embedFileMap;

    /**
     * 工作簿
     */
    private final Workbook workbook;

    public Excel2HtmlBuilder(byte[] fileData) throws IOException {
        this.fileData = fileData;
        this.sheetToHtmlMap = new HashMap<>();
        this.workbook = new XSSFWorkbook(new ByteArrayInputStream(fileData));
    }

    public Excel2HtmlBuilder(InputStream stream) throws IOException {
        this(FileUtil.getFileStream(stream));
    }

    public Excel2HtmlBuilder(File file) throws IOException {
        this(FileUtil.getFileStream(file));
    }

    public Excel2HtmlBuilder setCompressStyle(boolean compressStyle) {
        this.isCompressStyle = compressStyle;
        return this;
    }

    public Excel2HtmlBuilder setLoadEmbedFile(boolean loadEmbedFile) {
        this.isLoadEmbedFile = loadEmbedFile;
        return this;
    }

    public Excel2HtmlBuilder setTrimCellValue(boolean trimCellValue) {
        this.isTrimCellValue = trimCellValue;
        return this;
    }

    /**
     * 加载嵌入文件
     *
     * @throws IOException
     */
    private void doLoadEmbedFile() throws IOException {
        if (this.isLoadEmbedFile && this.embedFileMap == null) {
            // 解析嵌入图片数据
            Map<String, String> stringStringMap = CellEmbedFileParser.processZipEntries(new ByteArrayInputStream(fileData));
            this.embedFileMap = CellEmbedFileParser.processPictures(new ByteArrayInputStream(fileData), stringStringMap);
        }
    }

    /**
     * 获取最大列数
     *
     * @param sheet sheet
     */
    private int getMaxColNum(Sheet sheet) {
        short colNum = 0;
        for (Row row : sheet) {
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum > colNum) {
                colNum = lastCellNum;
            }
        }
        return colNum;
    }

    public HtmlPage buildHtml(int sheetIndex) throws IOException {
        if (this.workbook == null) {
            return null;
        }
        this.doLoadEmbedFile();
        HtmlPage htmlPage = this.sheetToHtmlMap.get(sheetIndex);
        if (Objects.nonNull(htmlPage)) {
            return htmlPage;
        }
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int maxColNum = this.getMaxColNum(sheet);
        htmlPage = Excel2Html.toHtml(sheet, maxColNum, this.isCompressStyle, this.isTrimCellValue, this.embedFileMap);
        sheetToHtmlMap.put(sheetIndex, htmlPage);
        return htmlPage;
    }

    public HtmlPage buildHtml(Sheet sheet) throws IOException {
        if (this.workbook == null || sheet == null || sheet.getWorkbook() != workbook) {
            return null;
        }
        int sheetIndex = this.workbook.getSheetIndex(sheet);
        return this.buildHtml(sheetIndex);
    }

    public List<HtmlPage> buildHtmlWithStartAndEndIndex(int startSheetIndex, int endSheetIndex) throws IOException {
        if (this.workbook == null) {
            return null;
        }
        this.doLoadEmbedFile();

        int endIndex = endSheetIndex;
        int numberOfSheets = this.workbook.getNumberOfSheets();
        if (endIndex > numberOfSheets - 1) {
            endIndex = numberOfSheets - 1;
        }

        List<HtmlPage> htmlList = new ArrayList<>();

        for (int i = startSheetIndex; i <= endIndex; i++) {
            HtmlPage htmlPage = this.buildHtml(i);
            htmlList.add(htmlPage);
        }
        return htmlList;
    }

    public List<HtmlPage> buildHtmlWithStartIndex(int startSheetIndex) throws IOException {
        return this.buildHtmlWithStartAndEndIndex(startSheetIndex, this.workbook.getNumberOfSheets() - 1);
    }
}
