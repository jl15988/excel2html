package com.jl15988.excel2html;

import com.jl15988.excel2html.converter.UnitConverter;
import com.jl15988.excel2html.converter.style.StyleConverter;
import com.jl15988.excel2html.converter.style.StyleGroupHtml;
import com.jl15988.excel2html.enums.ParserdCellValueType;
import com.jl15988.excel2html.html.HtmlElement;
import com.jl15988.excel2html.html.HtmlMeta;
import com.jl15988.excel2html.html.HtmlPage;
import com.jl15988.excel2html.html.IHtmlElement;
import com.jl15988.excel2html.model.parser.ParserdCellValue;
import com.jl15988.excel2html.model.parser.ParserdStyle;
import com.jl15988.excel2html.model.style.CommonCss;
import com.jl15988.excel2html.parser.CellStyleParser;
import com.jl15988.excel2html.parser.CellValueParser;
import com.jl15988.excel2html.parser.DrawingValueParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author Jalon
 * @since 2024/12/1 14:24
 **/
public class Excel2Html {

    /**
     * 表格转 html
     *
     * @param sheet         表格 sheet
     * @param columnNum     列数
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到 css
     */
    public static HtmlPage toHtml(Sheet sheet, int columnNum, boolean compressStyle, boolean trimCellValue) {
        return toHtml(sheet, columnNum, compressStyle, trimCellValue, null);
    }

    /**
     * 表格转 html
     *
     * @param sheet         表格 sheet
     * @param columnNum     列数
     * @param compressStyle 是否压缩样式，默认样式放在标签上，开启后会将大部分重复样式转到 css
     * @param embedFileMap  嵌入文件映射
     */
    public static HtmlPage toHtml(Sheet sheet, int columnNum, boolean compressStyle, boolean trimCellValue, Map<String, XSSFPictureData> embedFileMap) {
        HtmlPage htmlPage = getHtmlPage();
        HtmlElement div = new HtmlElement("div");
        div.addClass("exc-page");

        HtmlElement table = new HtmlElement("table");
        // 获取合并的单元格
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        // 用于开启样式压缩式缓存样式
        Map<String, Map<String, Object>> tagCellStyleCompressCache = new HashMap<>();
        Map<String, Map<String, Object>> tagCellContainerStyleCompressCache = new HashMap<>();
        Map<String, Map<String, Object>> tagCellValStyleCompressCache = new HashMap<>();

        // 单元格解析
        for (Row row : sheet) {
            HtmlElement tr = new HtmlElement("tr");
            for (int cellIndex = 0; cellIndex < columnNum; cellIndex++) {
                Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                ParserdCellValue parserdCellValue = CellValueParser.parseCellValue(cell, embedFileMap);
                String cellValue = parserdCellValue.getValue();
                if (trimCellValue) {
                    cellValue = cellValue.trim();
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

                ParserdStyle parserdStyle = CellStyleParser.parserCellStyle(cell, compressStyle);

                // 判断是否合并单元格，添加合并单元格属性
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
                                ParserdStyle mergedParserdStyle = CellStyleParser.parserCellStyle(lastColumnLastRowCell, compressStyle);

                                mergedParserdStyle.getCellStyle().forEach((name, value) -> {
                                    if (parserdStyle.hasCellStyle(name)) {
                                        if ((name.contains("-right") || name.contains("-bottom"))) {
                                            parserdStyle.addCellStyle(name, value);
                                        }
                                    } else {
                                        parserdStyle.addCellStyle(name, value);
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
                        String mergedTotalHeight = UnitConverter.convert().convertPointsString(totalHeight);
                        parserdStyle.addCellContainerStyle("height", mergedTotalHeight);
                        parserdStyle.addCellContainerStyle("max-height", mergedTotalHeight);
                        parserdStyle.addCellContainerStyle("min-height", mergedTotalHeight);
                    } else {
                        td.addClass("merged-display-cell");
                        // 忽略被合并的单元格
                        // todo 为了适配单元格内容显示与否，可能要追加并隐藏元素
//                        continue;
                    }
                }

                // 添加样式
                if (compressStyle) {
                    // 添加到tag-style map中，用于后面分组转换
                    tagCellStyleCompressCache.put(td.getUID(), parserdStyle.getCellStyle());
                } else {
                    td.setStyleMap(parserdStyle.getCellStyle());
                }

                HtmlElement cellContainerSpan = new HtmlElement("span")
                        .addClass("exc-table-cell-container");
                if (compressStyle) {
                    // 添加到tag-style map中，用于后面分组转换
                    tagCellContainerStyleCompressCache.put(cellContainerSpan.getUID(), parserdStyle.getCellContainerStyle());
                } else {
                    cellContainerSpan.setStyleMap(parserdStyle.getCellContainerStyle());
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
                    if (compressStyle) {
                        // 添加到tag-style map中，用于后面分组转换
                        cellValueSpan.addClasses(parserdStyle.getCellValClassList());
                        tagCellValStyleCompressCache.put(cellValueSpan.getUID(), parserdStyle.getCellValCellStyle());
                    } else {
                        cellValueSpan.setStyleMap(parserdStyle.getCellValCellStyle());
                    }
                    cellTableSpan.addChildElement(cellValueSpan);
                    cellContainerSpan.addChildElement(cellTableSpan);
                }

                td.addChildElement(cellContainerSpan);
                tr.addChildElement(td);
            }
            table.addChildElement(tr);
        }
        div.addChildElement(table);
        htmlPage.addElement(div);
        // 添加图片图形解析结果
        htmlPage.addElements(DrawingValueParser.parserDrawing(sheet));
        if (compressStyle) {
            // 添加通用样式
            htmlPage.addStyleContent(new CommonCss().toHtmlString());
            setCompressStyle(htmlPage, tagCellStyleCompressCache, tagCellContainerStyleCompressCache, tagCellValStyleCompressCache);
        }

        return htmlPage;
    }

    private static HtmlPage getHtmlPage() {
        HtmlPage htmlPage = new HtmlPage();
        htmlPage.setLang("zh-CN");
        htmlPage.addMeta(HtmlMeta
                        .builder()
                        .addAttr("charset", "UTF-8")
                        .build())
                .addStyleContent("* {\n" +
                        "            padding: 0;\n" +
                        "            margin: 0;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-page {\n" +
                        "            position: relative;\n" +
                        "        }\n" +
                        "\n" +
                        "        table {\n" +
                        "            table-layout: fixed;\n" +
                        "            box-sizing: border-box;\n" +
                        "            border-collapse: collapse;\n" +
                        "            border-spacing: 0;\n" +
                        "        }\n" +
                        "\n" +
                        "        td {\n" +
                        "            overflow: visible;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-table-cell.merged-cell {\n" +
                        "            overflow: hidden;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-table-cell.merged-display-cell {\n" +
                        "            display: none;\n" +
                        "        }\n" +
                        "\n" +
                        // 原本container直接包含图片数据，但是无法使用背景色覆盖前者，现又添加了一层img-container
                        "        .exc-table-cell.embed-img-data .exc-table-cell-container {\n" +
//                        "            display: block;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-table-cell.has-data .exc-table-cell-table {\n" +
                        "            background-color: white;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-table-cell + .has-data {\n" +
                        // 本来通过每个单元格添加是否有值来限制兄弟节点超出隐藏，但是不能夸单元格，然后改为了通过有数据添加背景色来覆盖前者数据
//                        "            overflow: hidden;\n" +
                        "        }\n" +
                        "\n" +
                        "        .exc-table-cell-container {\n" +
                        "            display: flex;\n" +
                        "            width: 100%;\n" +
                        "        }\n" +
                        "        \n" +
                        "        .exc-table-cell-table {\n" +
                        "            display: table;\n" +
                        "            width: 100%;\n" +
                        "        }\n" +
                        "        \n" +
                        "        .exc-table-val {\n" +
                        "            display: table-cell;\n" +
                        "            padding-top: 2px;\n" +
                        "        }" +
                        "        \n" +
                        "        .embed-img-container {\n" +
                        "            display: block;\n" +
                        "            width: 100%;\n" +
                        "            height: 100%;\n" +
                        "            background-color: white;\n" +
                        "        }" +
                        "        \n" +
                        "        .embed_img {\n" +
                        "            width: 100%;\n" +
                        "            height: 100%;\n" +
                        "            object-fit: contain;\n" +
                        "        }");
        return htmlPage;
    }

    private static void setCompressStyle(HtmlPage htmlPage,
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
    private static void addTagStyleClass(IHtmlElement<?> htmlElement, Map<String, List<String>> tagStyleUidMap) {
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
