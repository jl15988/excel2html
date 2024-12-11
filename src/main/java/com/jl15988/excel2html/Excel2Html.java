package com.jl15988.excel2html;

import com.jl15988.excel2html.converter.style.StyleConverter;
import com.jl15988.excel2html.converter.style.StyleGroupHtml;
import com.jl15988.excel2html.enums.ParserdCellValueType;
import com.jl15988.excel2html.formatter.DefaultCellValueFormatter;
import com.jl15988.excel2html.formatter.ICellValueFormater;
import com.jl15988.excel2html.handler.ICellHandler;
import com.jl15988.excel2html.handler.ITrElementHandler;
import com.jl15988.excel2html.html.HtmlElement;
import com.jl15988.excel2html.html.HtmlMeta;
import com.jl15988.excel2html.html.HtmlPage;
import com.jl15988.excel2html.html.IHtmlElement;
import com.jl15988.excel2html.model.parser.ParserdCellValue;
import com.jl15988.excel2html.model.parser.ParserdStyleResult;
import com.jl15988.excel2html.model.style.CommonCss;
import com.jl15988.excel2html.model.unit.Pixel;
import com.jl15988.excel2html.model.unit.Point;
import com.jl15988.excel2html.parser.CellStyleParser;
import com.jl15988.excel2html.parser.CellValueParser;
import com.jl15988.excel2html.parser.DrawingValueParser;
import com.jl15988.excel2html.utils.FileUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
public class Excel2Html {

    /**
     * dpi
     */
    private int dpi = 96;

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
    private final Map<Integer, HtmlPage> sheetToHtmlMap;

    /**
     * 是否启用压缩样式
     */
    private boolean isCompressStyle = true;

    /**
     * 是否加载嵌入文件
     */
    private boolean isLoadEmbedFile = true;

    /**
     * 嵌入文件缓存
     */
    private Map<String, XSSFPictureData> embedFileMap;

    /**
     * 工作簿
     */
    private Workbook workbook;

    /**
     * 单元格格式化
     */
    private ICellValueFormater cellValueFormater;

    /**
     * tr 元素处理器
     */
    private ITrElementHandler trElementHandler;

    /**
     * 单元格处理器
     */
    private ICellHandler cellHandler;

    public Excel2Html(byte[] fileData) throws IOException {
        this.fileData = fileData;
        this.sheetToHtmlMap = new HashMap<>();
        this.cellValueFormater = new DefaultCellValueFormatter();
        this.workbook = new XSSFWorkbook(new ByteArrayInputStream(fileData));
    }

    public Excel2Html(InputStream stream) throws IOException {
        this(FileUtil.getFileStream(stream));
    }

    public Excel2Html(File file) throws IOException {
        this(FileUtil.getFileStream(file));
    }

    public Excel2Html() {
        this.sheetToHtmlMap = new HashMap<>();
        this.cellValueFormater = new DefaultCellValueFormatter();
    }

    public Excel2Html setDpi(int dpi) {
        this.dpi = dpi;
        return this;
    }

    public Excel2Html setCompressStyle(boolean compressStyle) {
        this.isCompressStyle = compressStyle;
        return this;
    }

    public Excel2Html setLoadEmbedFile(boolean loadEmbedFile) {
        this.isLoadEmbedFile = loadEmbedFile;
        return this;
    }

    public Excel2Html setEmbedFileMap(Map<String, XSSFPictureData> embedFileMap) {
        this.embedFileMap = embedFileMap;
        return this;
    }

    public Excel2Html setCellValueFormater(ICellValueFormater cellValueFormater) {
        this.cellValueFormater = cellValueFormater;
        return this;
    }

    public Excel2Html setTrElementHandler(ITrElementHandler trElementHandler) {
        this.trElementHandler = trElementHandler;
        return this;
    }

    public Excel2Html setCellHandler(ICellHandler cellHandler) {
        this.cellHandler = cellHandler;
        return this;
    }

    /**
     * 加载嵌入文件
     *
     * @throws IOException
     */
    private void doLoadEmbedFile() throws IOException {
        if (Objects.nonNull(fileData) && this.isLoadEmbedFile && this.embedFileMap == null) {
            this.embedFileMap = Excel2HtmlUtil.doLoadEmbedFile(fileData);
        }
    }

    /**
     * 构建 html
     *
     * @param sheet         sheet
     * @param startRowIndex 开始行，可为空默认0
     * @param endRowIndex   结束行，可为空默认最后
     * @param startColIndex 开始列，可为空默认0
     * @param endColIndex   结束列，可为空默认最后
     * @return html 结果
     * @throws IOException
     */
    public HtmlPage buildHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex) throws IOException {
        if (Objects.isNull(sheet)) return null;
        // 加载嵌入文件
        this.doLoadEmbedFile();

        // 从 sheet 中获取工作簿
        if (Objects.isNull(this.workbook)) {
            this.workbook = sheet.getWorkbook();
        }

        int sheetIndex = -1;
        // 尝试从缓存中获取
        if (Objects.nonNull(this.workbook)) {
            if (sheet.getWorkbook() == this.workbook) {
                sheetIndex = this.workbook.getSheetIndex(sheet);
                if (sheetIndex != -1) {
                    HtmlPage htmlPage = this.sheetToHtmlMap.get(sheetIndex);
                    if (Objects.nonNull(htmlPage)) {
                        return htmlPage;
                    }
                }
            }
        }

        int rowEndIndex = Optional.ofNullable(endRowIndex).orElse(Excel2HtmlUtil.getMaxRowNum(sheet) - 1);
        if (rowEndIndex < 0) {
            return null;
        }
        // 如果没有指定单元格数量则取最大单元格数
        int colEndIndex = Optional.ofNullable(endColIndex).orElse(Excel2HtmlUtil.getMaxColNum(sheet) - 1);
        if (colEndIndex < 0) {
            return null;
        }

        HtmlPage htmlPage = this.doBuildHtml(sheet, startRowIndex, rowEndIndex, startColIndex, colEndIndex);
        // 缓存结果
        if (sheetIndex != -1) {
            this.sheetToHtmlMap.put(sheetIndex, htmlPage);
        }
        return htmlPage;
    }

    /**
     * 通过 sheet 索引构建 html
     *
     * @param sheetIndex    sheet 下标
     * @param startRowIndex 开始行，可为空默认0
     * @param endRowIndex   结束行，可为空默认最后
     * @param startColIndex 开始列，可为空默认0
     * @param endColIndex   结束列，可为空默认最后
     * @return html 结果
     * @throws IOException
     */
    public HtmlPage buildHtmlWithSheetIndex(int sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex) throws IOException {
        if (this.workbook == null) {
            return null;
        }
        Sheet sheet = this.workbook.getSheetAt(sheetIndex);
        return this.buildHtml(sheet, startRowIndex, endRowIndex, startColIndex, endColIndex);
    }

    /**
     * 通过 sheet 索引构建 html
     *
     * @param startSheetIndex sheet 开始下标，可为空默认0
     * @param endSheetIndex   sheet 结束下标，可为空默认最后
     * @param startRowIndex   开始行，可为空默认0
     * @param endRowIndex     结束行，可为空默认最后
     * @param startColIndex   开始列，可为空默认0
     * @param endColIndex     结束列，可为空默认最后
     * @return html 结果
     * @throws IOException
     */
    public List<HtmlPage> buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex) throws IOException {
        if (this.workbook == null) {
            return null;
        }
        this.doLoadEmbedFile();

        int numberOfSheets = this.workbook.getNumberOfSheets();
        int endIndex = Optional.ofNullable(endSheetIndex).orElse(numberOfSheets - 1);
        if (endIndex > numberOfSheets - 1) {
            endIndex = numberOfSheets - 1;
        }

        List<HtmlPage> htmlList = new ArrayList<>();

        for (int i = Optional.ofNullable(startSheetIndex).orElse(0); i <= endIndex; i++) {
            HtmlPage htmlPage = this.buildHtmlWithSheetIndex(i, startRowIndex, endRowIndex, startColIndex, endColIndex);
            htmlList.add(htmlPage);
        }
        return htmlList;
    }

    /**
     * 构建 html
     *
     * @param sheet sheet
     * @return html 结果
     * @throws IOException
     */
    public HtmlPage buildHtml(Sheet sheet) throws IOException {
        return this.buildHtml(sheet, null, null, null, null);
    }

    /**
     * 通过 sheet 索引构建 html
     *
     * @param sheetIndex sheet 下标
     * @return html 结果
     * @throws IOException
     */
    public HtmlPage buildHtmlWithSheetIndex(int sheetIndex) throws IOException {
        return this.buildHtmlWithSheetIndex(sheetIndex, null, null, null, null);
    }

    /**
     * 通过 sheet 索引构建 html
     *
     * @param startSheetIndex sheet 开始下标，可为空默认0
     * @param endSheetIndex   sheet 结束下标，可为空默认最后
     * @return html 结果
     * @throws IOException
     */
    public List<HtmlPage> buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex) throws IOException {
        return this.buildHtmlWithSheetIndex(startSheetIndex, endSheetIndex, null, null, null, null);
    }

    private HtmlPage doBuildHtml(Sheet sheet, Integer startRowIndex, int endRowIndex, Integer startColIndex, int endColIndex) {
        // 赋值默认列宽
        sheet.setDefaultColumnWidth(Excel2HtmlUtil.getDefaultColumnWidthSpecial(workbook));

        HtmlPage htmlPage = getHtmlPage();
        HtmlElement div = new HtmlElement("div");
        div.addClass("exc-page");

        HtmlElement table = new HtmlElement("table");
        table.addAttribute("border", "0");
        table.addAttribute("cellpadding", "0");
        table.addAttribute("cellspacing", "0");
        // 获取合并的单元格
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        // 用于开启样式压缩式缓存样式
        Map<String, Map<String, Object>> tagCellStyleCompressCache = new HashMap<>();
        Map<String, Map<String, Object>> tagCellContainerStyleCompressCache = new HashMap<>();
        Map<String, Map<String, Object>> tagCellValStyleCompressCache = new HashMap<>();

        // todo 按打印页转换

        // 单元格解析
        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        int defaultColumnWidthInPixels = Excel2HtmlUtil.getDefaultColumnWidthInPixels(workbook);
        for (int rowIndex = Optional.ofNullable(startRowIndex).orElse(0); rowIndex <= endRowIndex; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            HtmlElement tr = new HtmlElement("tr");

            Integer startCol = Optional.ofNullable(startColIndex).orElse(0);
            if (Objects.isNull(row)) {
                // 对于为空的行，添加默认的单元格
                for (int cellIndex = startCol; cellIndex <= endColIndex; cellIndex++) {
                    HtmlElement td = new HtmlElement("td");
                    td.addStyle("height", new Point(defaultRowHeightInPoints, this.dpi).toString());
                    td.addStyle("width", new Pixel(defaultColumnWidthInPixels, this.dpi).toString());
                    tr.addChildElement(td);
                }

                table.addChildElement(tr);
                continue;
            }

            for (int cellIndex = startCol; cellIndex <= endColIndex; cellIndex++) {
                Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                ParserdCellValue parserdCellValue = CellValueParser.parseCellValue(cell, this.embedFileMap);
                String cellValue = parserdCellValue.getValue();

                // 单元格内容格式化
                if (Objects.nonNull(cellValueFormater)) {
                    cellValue = cellValueFormater.format(cellValue, cell);
                }

                boolean valueEmpty = cellValue == null || cellValue.isEmpty();

                HtmlElement td = new HtmlElement("td");
                td.addClass("exc-table-cell");
                // 根据单元格是否有值，添加 class
                if (valueEmpty) {
                    td.addClass("no-data");
                } else {
                    td.addClass("has-data");
                }

                // 解析单元格样式
                ParserdStyleResult parserdStyleResult = CellStyleParser.parserCellStyle(cell, this.dpi);
                td.addClasses(parserdStyleResult.getCellClassList());

                // 解析合并单元格
                parserMergedCell(mergedRegions, cell, td, parserdStyleResult);

                // 执行单元格处理器
                if (Objects.nonNull(this.cellHandler)) {
                    this.cellHandler.handle(td, rowIndex, cellIndex, sheet);
                    this.cellHandler.handleStyle(parserdStyleResult, cell, rowIndex, cellIndex);
                }

                // 添加样式
                Map<String, Object> cellStyleMap = parserdStyleResult.getCellStyle();
                if (cellStyleMap.containsKey("background-color")) {
                    td.addClass("has-bg-color");
                }
                if (this.isCompressStyle) {
                    // 添加到tag-style map中，用于后面分组转换
                    tagCellStyleCompressCache.put(td.getUID(), cellStyleMap);
                } else {
                    td.setStyleMap(cellStyleMap);
                }

                HtmlElement cellContainerSpan = new HtmlElement("span")
                        .addClass("exc-table-cell-container");
                if (this.isCompressStyle) {
                    // 添加到tag-style map中，用于后面分组转换
                    tagCellContainerStyleCompressCache.put(cellContainerSpan.getUID(), parserdStyleResult.getCellContainerStyle());
                } else {
                    cellContainerSpan.setStyleMap(parserdStyleResult.getCellContainerStyle());
                }

                if (ParserdCellValueType.HTML_IMG.equals(parserdCellValue.getType())) {
                    // 嵌入图片的特殊处理
                    td.addClass("embed-img-data");
                    cellContainerSpan.addChildElement(HtmlElement.builder("span")
                            .addClass("embed-img-container")
                            .content(cellValue)
                            .build());
                } else {
                    HtmlElement cellTableSpan = new HtmlElement("span")
                            .addClass("exc-table-cell-table");
                    HtmlElement cellValueSpan = new HtmlElement("span")
                            .addClass("exc-table-val")
                            .setContent(cellValue);
                    if (this.isCompressStyle) {
                        // 添加到tag-style map中，用于后面分组转换
                        cellValueSpan.addClasses(parserdStyleResult.getCellValStyleClassList());
                        tagCellValStyleCompressCache.put(cellValueSpan.getUID(), parserdStyleResult.getCellValCellStyle());
                    } else {
                        cellValueSpan.setStyleMap(parserdStyleResult.getCellValCellStyle());
                    }
                    cellTableSpan.addChildElement(cellValueSpan);
                    cellContainerSpan.addChildElement(cellTableSpan);
                }

                td.addChildElement(cellContainerSpan);
                tr.addChildElement(td);
            }

            // 执行 tr 元素处理器
            if (Objects.nonNull(this.trElementHandler)) {
                tr = this.trElementHandler.handle(tr, rowIndex, sheet);
            }

            table.addChildElement(tr);
        }
        div.addChildElement(table);
        htmlPage.addElement(div);
        // 添加图片图形解析结果
        htmlPage.addElements(DrawingValueParser.parserDrawing(sheet));
        if (this.isCompressStyle) {
            // 添加通用样式
            htmlPage.addStyleContent(new CommonCss().toHtmlString());
            setCompressStyle(htmlPage, tagCellStyleCompressCache, tagCellContainerStyleCompressCache, tagCellValStyleCompressCache);
        }

        return htmlPage;
    }

    /**
     * 解析合并单元格
     */
    private void parserMergedCell(List<CellRangeAddress> mergedRegions, Cell cell, HtmlElement td, ParserdStyleResult parserdStyleResult) {
        // 判断是否合并单元格，添加合并单元格属性
        Sheet sheet = cell.getRow().getSheet();
        CellRangeAddress cellAddresses = mergedRegions.stream().filter(address -> address.isInRange(cell)).findFirst().orElse(null);
        if (Objects.nonNull(cellAddresses)) {
            if (cellAddresses.getFirstRow() == cell.getRowIndex() && cellAddresses.getFirstColumn() == cell.getColumnIndex()) {
                td.addClass("merged-cell");
                // 对合并单元格的第一行第一列单元格处理
                int rowSpan = cellAddresses.getLastRow() - cellAddresses.getFirstRow() + 1;
                int colSpan = cellAddresses.getLastColumn() - cellAddresses.getFirstColumn() + 1;
                if (rowSpan > 1) {
                    td.addAttribute("rowspan", String.valueOf(rowSpan));
                }
                if (colSpan > 1) {
                    td.addAttribute("colspan", String.valueOf(colSpan));
                }

                // 合并单元格样式
                // 取最后一行最后一个单元格样式
                int lastRowIndex = cellAddresses.getLastRow();
                int lastColumnIndex = cellAddresses.getLastColumn();
                Row lastRow = sheet.getRow(lastRowIndex);
                if (Objects.nonNull(lastRow)) {
                    Cell lastColumnLastRowCell = lastRow.getCell(lastColumnIndex);
                    if (Objects.nonNull(lastColumnLastRowCell)) {
                        ParserdStyleResult mergedParserdStyleResult = CellStyleParser.parserCellStyle(lastColumnLastRowCell, this.dpi);

                        mergedParserdStyleResult.getCellStyle().forEach((name, value) -> {
                            if (parserdStyleResult.hasCellStyle(name)) {
                                if ((name.contains("-right") || name.contains("-bottom"))) {
                                    parserdStyleResult.addCellStyle(name, value);
                                }
                            } else {
                                parserdStyleResult.addCellStyle(name, value);
                            }
                        });
                    }
                }

                double totalHeight = 0;
                int firstRowIndex = cellAddresses.getFirstRow();
                for (int j = firstRowIndex; j <= lastRowIndex; j++) {
                    Row rowItem = sheet.getRow(j);
                    if (Objects.nonNull(rowItem)) {
                        totalHeight += rowItem.getHeightInPoints();
                    }
                }
                String mergedTotalHeightC = new Point(totalHeight - new Pixel(3, dpi).toPoint().getValue(), this.dpi).toString();
                String mergedTotalHeight = new Point(totalHeight, this.dpi).toString();
                parserdStyleResult.addCellContainerStyle("height", mergedTotalHeightC);
                parserdStyleResult.addCellContainerStyle("max-height", mergedTotalHeightC);
                parserdStyleResult.addCellContainerStyle("min-height", mergedTotalHeightC);
                parserdStyleResult.addCellStyle("height", mergedTotalHeight);
            } else {
                td.addClass("merged-display-cell");
                // 忽略被合并的单元格
                // todo 为了适配单元格内容显示与否，可能要追加并隐藏元素
//                        continue;
            }
        }
    }

    private HtmlPage getHtmlPage() {
        HtmlPage htmlPage = new HtmlPage();
        htmlPage.setLang("zh-CN");
        htmlPage.addMeta(HtmlMeta
                        .builder()
                        .addAttr("charset", "UTF-8")
                        .build())
                .addStyleContent("" +
//                        "* {\n" +
//                        "            padding: 0;\n" +
//                        "            margin: 0;\n" +
//                        "        }\n" +
//                        "\n" +
                        // 基础样式
                        ".exc-page {\n" +
                        "    position: relative;\n" +
                        "}\n" +
                        "table {\n" +
                        "    table-layout: fixed;\n" +
                        "    box-sizing: border-box;\n" +
                        "    border-collapse: collapse;\n" +
                        "    border-spacing: 0;\n" +
                        "}\n" +
                        "td {\n" +
                        "    overflow: visible;\n" +
                        "    box-sizing: border-box;\n" +
                        "    mso-style-parent: style0;\n" +
                        "    padding-top: 1px;\n" +
                        "    padding-right: 1px;\n" +
                        "    padding-left: 1px;\n" +
                        "    mso-ignore: padding;\n" +
                        "    mso-number-format: \"General\";\n" +
                        "    text-align: general;\n" +
                        "    vertical-align: middle;\n" +
                        "    white-space: nowrap;\n" +
                        "    mso-rotate: 0;\n" +
                        "    mso-pattern: auto;\n" +
                        "    mso-background-source: auto;\n" +
                        "    color: #000000;\n" +
                        "    font-size: 11.0pt;\n" +
                        "    font-weight: 400;\n" +
                        "    font-style: normal;\n" +
                        "    text-decoration: none;\n" +
                        "    font-family: 宋体;\n" +
                        "    mso-generic-font-family: auto;\n" +
                        "    mso-font-charset: 134;\n" +
                        "    border: none;\n" +
                        "    mso-protection: locked visible;" +
                        "}\n" +
                        // 合并的单元格超出隐藏
                        ".exc-table-cell.merged-cell {\n" +
                        "    overflow: hidden;\n" +
                        "}\n" +
                        // 合并隐藏的隐藏
                        ".exc-table-cell.merged-display-cell {\n" +
                        "    display: none;\n" +
                        "}\n" +
                        // 原本container直接包含图片数据，但是无法使用背景色覆盖前者，现又添加了一层img-container
                        ".exc-table-cell.embed-img-data .exc-table-cell-container {\n" +
//                        "    display: block;\n" +
                        "}\n" +
                        // 有数据的单元格添加背景色
                        ".exc-table-cell.has-data .exc-table-cell-table {\n" +
                        "    background-color: white;\n" +
                        "}\n" +
                        // 有背景的单元格背景调整背景色
                        ".exc-table-cell.has-bg-color .exc-table-cell-table {\n" +
                        "    background-color: rgba(0, 0, 0, 0);\n" +
                        "}\n" +
                        ".exc-table-cell.has-bg-color .embed-img-container {\n" +
                        "    background-color: rgba(0, 0, 0, 0);\n" +
                        "}\n" +
                        // 本来通过每个单元格添加是否有值来限制兄弟节点超出隐藏，但是不能夸单元格，然后改为了通过有数据添加背景色来覆盖前者数据
                        ".exc-table-cell + .has-data {\n" +
//                        "    overflow: hidden;\n" +
                        "}\n" +
                        // 下面的元素、table、value容器是为了尽可能的还原 excel 单元格展示样式
                        // 单元格内元素容器
                        ".exc-table-cell-container {\n" +
                        "    display: flex;\n" +
                        "    width: 100%;\n" +
                        "}\n" +
                        // 单元格内 table 容器
                        ".exc-table-cell-table {\n" +
                        "    display: table;\n" +
                        "    width: 100%;\n" +
                        "}\n" +
                        // 单元格内 value 容器
                        ".exc-table-val {\n" +
                        "    display: table-cell;\n" +
//                        "    padding-top: 2px;\n" +
                        "}" +
                        // 单元格内图片容器，为了还原嵌入图片样式
                        ".embed-img-container {\n" +
                        "    display: block;\n" +
                        "    width: 100%;\n" +
                        "    height: 100%;\n" +
                        "    background-color: white;\n" +
                        "}" +
                        // 嵌入图片
                        ".embed_img {\n" +
                        "    width: 100%;\n" +
                        "    height: 100%;\n" +
                        "    object-fit: contain;\n" +
                        "}" +
                        // 换行的单元格尾部连续空格不占用空间
                        ".exc-table-cell.wrap-cell .value-end-spaces {\n" +
                        "    white-space: normal;\n" +
                        "}");
        return htmlPage;
    }

    private void setCompressStyle(HtmlPage htmlPage,
                                  Map<String, Map<String, Object>> tagCellStyleCompressCache,
                                  Map<String, Map<String, Object>> tagCellContainerStyleCompressCache,
                                  Map<String, Map<String, Object>> tagCellValStyleCompressCache) {
        // 对 style 进行解析分组，转 class，压缩
        StyleGroupHtml cellStyleHtml = StyleConverter.tagStyleToHtmlString(tagCellStyleCompressCache);
        StyleGroupHtml cellContainerStyleHtml = StyleConverter.tagStyleToHtmlString(tagCellContainerStyleCompressCache);
        StyleGroupHtml cellValStyleHtml = StyleConverter.tagStyleToHtmlString(tagCellValStyleCompressCache);

        // 添加 style 内容
        htmlPage.addStyleContent(cellStyleHtml.getStyleContent());
        htmlPage.addStyleContent(cellContainerStyleHtml.getStyleContent());
        htmlPage.addStyleContent(cellValStyleHtml.getStyleContent());

        Map<String, List<String>> tagStyleUidMap = new HashMap<>();
        tagStyleUidMap.putAll(cellStyleHtml.getTagStyleUidMap());
        tagStyleUidMap.putAll(cellContainerStyleHtml.getTagStyleUidMap());
        tagStyleUidMap.putAll(cellValStyleHtml.getTagStyleUidMap());

        // 匹配每个元素的 class
        for (HtmlElement htmlElement : htmlPage.getElementList()) {
            addTagStyleClass(htmlElement, tagStyleUidMap);
        }
    }

    /**
     * 给元素添加样式 class
     *
     * @param htmlElement    元素
     * @param tagStyleUidMap 元素-样式id
     */
    private void addTagStyleClass(IHtmlElement<?> htmlElement, Map<String, List<String>> tagStyleUidMap) {
        if (null == tagStyleUidMap || tagStyleUidMap.isEmpty()) return;
        List<IHtmlElement<?>> childrenElementList = htmlElement.getChildrenElementList();
        if (null != childrenElementList && !childrenElementList.isEmpty()) {
            for (IHtmlElement<?> child : childrenElementList) {
                addTagStyleClass(child, tagStyleUidMap);
            }
        }
        if (tagStyleUidMap.containsKey(htmlElement.getUID())) {
            htmlElement.addClasses(tagStyleUidMap.get(htmlElement.getUID()));
        }
    }
}
