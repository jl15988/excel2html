package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.Excel2HtmlUtil;
import com.jl15988.excel2html.enums.CommonElementClass;
import com.jl15988.excel2html.enums.ParserdCellValueType;
import com.jl15988.excel2html.html.HtmlElement;
import com.jl15988.excel2html.html.HtmlElementList;
import com.jl15988.excel2html.model.parser.ParserdCellValue;
import com.jl15988.excel2html.model.style.FontICssStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.Base64;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 单元格内容解析器
 *
 * @author Jalon
 * @since 2024/11/29 16:50
 **/
public class CellValueParser {

    /**
     * 判断单元格内是否为富文本
     *
     * @param cell 单元格
     * @return 是否为富文本
     */
    public static boolean isRichValue(Cell cell) {
        CellType cellType = cell.getCellType();
        if (!CellType.STRING.equals(cellType)) {
            return false;
        }
        XSSFRichTextString richText = (XSSFRichTextString) cell.getRichStringCellValue();
        if (richText == null) {
            return false;
        }
        int formattingRuns = richText.numFormattingRuns();
        return formattingRuns > 0;
    }

    /**
     * 解析富文本单元格
     *
     * @param cell 单元格
     * @return 解析后的html列表
     */
    public static HtmlElementList parserCellRichValue(Cell cell) {
        if (!isRichValue(cell)) {
            return null;
        }

        // 单元格内容是否自动换行
        boolean wrapText = cell.getCellStyle().getWrapText();

        HtmlElementList htmlElementList = new HtmlElementList();
        // 获取富文本内容
        XSSFRichTextString richTextString = (XSSFRichTextString) cell.getRichStringCellValue();
        String richText = richTextString.toString();

        // 判断是否有尾部空白内容
        Matcher matcher = Pattern.compile("\\s+$").matcher(richText);
        boolean hasEmpty = matcher.find();
        // 空白开始下标
        Integer emptyStart = null;
        if (hasEmpty) {
            emptyStart = matcher.start();
        }

        int formattingRuns = richTextString.numFormattingRuns();
        for (int i = 0; i < formattingRuns; i++) {
            int start = richTextString.getIndexOfFormattingRun(i);
            int length = richTextString.getLengthOfFormattingRun(i);
            XSSFFont xssfFont = richTextString.getFontOfFormattingRun(i);
            String text = richText.substring(start, start + length);

            // 解析字体样式
            FontICssStyle fontCssStyle = XSSFFontParser.parserXSSFFontToStyleMap(xssfFont);

            if (wrapText && hasEmpty) {
                // 判断内容尾部是否有空白字符，如果有，且之后全是开白字符，则使之不占用空间
                if (emptyStart >= start && emptyStart <= (start + length)) {
                    htmlElementList.add(HtmlElement.builder("span")
                            .addStyle(fontCssStyle)
                            .addChildElements(
                                    HtmlElement.builder("")
                                            .content(richText.substring(start, emptyStart))
                                            .build(),
                                    HtmlElement.builder("span")
                                            .addStyle(fontCssStyle)
                                            .addClass(CommonElementClass.VALUE_END_SPACES.value())
                                            .content(text.substring(emptyStart - start))
                                            .build()
                            ).build()
                    );
                    continue;
                } else if (start > emptyStart) {
                    htmlElementList.add(HtmlElement.builder("span")
                            .addStyle(fontCssStyle)
                            .addClass(CommonElementClass.VALUE_END_SPACES.value())
                            .content(text)
                            .build()
                    );
                    continue;
                }
            }

            // 构建 span 标签
            HtmlElement spanEle = new HtmlElement("span");
            spanEle.addStyle(fontCssStyle).setContent(text);
            htmlElementList.add(spanEle);
        }
        return htmlElementList;
    }

    /**
     * 解析单元格数字内容
     *
     * @param cell     单元格
     * @param cellType 单元格类型
     * @return 数字内容
     */
    public static String parserCellNumericValue(Cell cell, CellType cellType) {
        if (cellType != CellType.NUMERIC) {
            return null;
        }

        // 判断是否为日期类型
        if (Excel2HtmlUtil.isCellDateFormatted(cell)) {
            // 使用增强版的格式获取方法
            String enhancedDataFormat = Excel2HtmlUtil.getDataFormatString(cell);
            // 如果是日期类型，并且我们有增强的格式
            if (enhancedDataFormat != null && !enhancedDataFormat.isEmpty()) {
                return CellValueFormatter.formatDateValue(cell, enhancedDataFormat);
            } else {
                // 没有增强格式或格式无效，使用DataFormatter
                DataFormatter dataFormatter = new DataFormatter();
                dataFormatter.setUseCachedValuesForFormulaCells(true);
                return dataFormatter.formatCellValue(cell);
            }
        }

        return CellValueFormatter.formatNumericValue(cell, cell.getNumericCellValue());
    }

    /**
     * 解析公式内容
     *
     * @param cell 单元格
     * @return 公式结果
     */
    public static String parserCellFormulaValue(Cell cell) {
        // 获取缓存公式结果类型
        CellType cellType = cell.getCachedFormulaResultType();
        return parseCellBaseValue(cell, cellType);
    }

    /**
     * 执行单元格公式，某个函数可能不支持
     *
     * @param cell 单元格
     * @return 执行公式后的结果
     */
    public static String exeCellFormula(Cell cell) {
        String resultValue = "";
        Workbook workbook = cell.getRow().getSheet().getWorkbook();
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        // 执行公式
        CellValue cellValue = formulaEvaluator.evaluate(cell);
        CellType cellType = cellValue.getCellType();
        try {
            switch (cellType) {
                case STRING:
                    resultValue = cellValue.getStringValue();
                    break;
                case NUMERIC:
                    resultValue = CellValueFormatter.formatNumericValue(cell, cellValue.getNumberValue());
                    break;
                case BOOLEAN:
                    resultValue = String.valueOf(cellValue.getBooleanValue()).toUpperCase(Locale.ROOT);
                    break;
                case BLANK:
                    resultValue = "";
                    break;
                default:
                    resultValue = "";
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return resultValue;
    }

    /**
     * 解析基础单元内容（不包含公式）
     *
     * @param cell     单元格
     * @param cellType 单元格类型
     * @return 解析后的内容
     */
    public static String parseCellBaseValue(Cell cell, CellType cellType) {
        String resultValue = "";
        switch (cellType) {
            case STRING:
                HtmlElementList htmlElementList = parserCellRichValue(cell);
                if (Objects.isNull(htmlElementList)) {
                    resultValue = cell.getStringCellValue();
                } else {
                    resultValue = htmlElementList.toHtmlString();
                }
                break;
            case NUMERIC:
                resultValue = parserCellNumericValue(cell, cellType);
                break;
            case BOOLEAN:
                resultValue = String.valueOf(cell.getBooleanCellValue()).toUpperCase(Locale.ROOT);
                break;
            case BLANK:
                resultValue = "";
                break;
            default:
                resultValue = "";
                break;
        }
        return resultValue;
    }

    /**
     * 解析单元格嵌入附件
     *
     * @param cell         单元格
     * @param embedFileMap 嵌入附件映射
     * @return 解析后的单元格内容
     */
    public static String parserCellEmbedFile(Cell cell, Map<String, XSSFPictureData> embedFileMap) {
        if (Objects.isNull(embedFileMap)) return null;
        String regex = "_xlfn.DISPIMG\\(\"([^)]+)\",\\d+\\)";
        Pattern pattern = Pattern.compile(regex);
        // 创建Matcher对象
        Matcher matcher = pattern.matcher(cell.getCellFormula());

        // 查找匹配项
        if (matcher.find()) {
            // 提取并打印ID
            String id = matcher.group(1);
            XSSFPictureData pictureData = embedFileMap.get(id);
            if (Objects.isNull(pictureData)) {
                return null;
            }
            byte[] imageBytes = pictureData.getData();
            String base64Image = "data:image/png;base64," + Base64.getEncoder().encodeToString(imageBytes);
            HtmlElement img = new HtmlElement("img");
            img.addClass("embed_img embed_img_" + id);
            img.addAttribute("src", base64Image);
            return img.toHtmlString();
        } else {
            return null;
        }
    }

    /**
     * 解析单元格内容
     *
     * @param cell 单元格
     * @return 解析后的单元格内容
     */
    public static ParserdCellValue parseCellValue(Cell cell, Map<String, XSSFPictureData> embedFileMap) {
        ParserdCellValue.ParserdCellValueBuilder parserdCellValueBuilder = ParserdCellValue.builder().type(ParserdCellValueType.TEXT);

        CellType cellType = cell.getCellType();
        if (CellType.STRING.equals(cellType)) {
            // 对于 string 类型，需要解析富文本
            HtmlElementList htmlElementList = parserCellRichValue(cell);
            if (Objects.isNull(htmlElementList)) {
                String stringCellValue = cell.getStringCellValue();

                if (cell.getCellStyle().getWrapText()) {
                    // 判断单元格内容尾部是否有空白字符串，有的话处理成不占用空间的元素
                    Matcher matcher = Pattern.compile("\\s+$").matcher(stringCellValue);
                    if (matcher.find()) {
                        int emptyStart = matcher.start();
                        stringCellValue = matcher.replaceAll("") + HtmlElement.builder("span")
                                .addClass(CommonElementClass.VALUE_END_SPACES.value())
                                .content(stringCellValue.substring(emptyStart))
                                .build().toHtmlString();
                    }
                }
                return parserdCellValueBuilder.value(stringCellValue).build();
            } else {
                return parserdCellValueBuilder.type(ParserdCellValueType.RICH_HTML_CONTENT).value(htmlElementList.toHtmlString()).build();
            }
        } else if (Objects.nonNull(embedFileMap) && CellType.FORMULA.equals(cellType)) {
            String cellFormula = cell.getCellFormula();
            if (cellFormula.startsWith("_xlfn.DISPIMG(\"")) {
                // 处理嵌入图片
                String embedValue = parserCellEmbedFile(cell, embedFileMap);
                if (Objects.nonNull(embedValue))
                    return parserdCellValueBuilder.type(ParserdCellValueType.HTML_IMG).value(embedValue).build();
            }
        }

        String value = "";

        // 判断是否为日期类型
        if (cellType == CellType.NUMERIC && Excel2HtmlUtil.isCellDateFormatted(cell)) {
            value = CellValueParser.parserCellNumericValue(cell, cellType);
        } else {
            // 非日期类型，使用DataFormatter
            DataFormatter dataFormatter = new DataFormatter();
            dataFormatter.setUseCachedValuesForFormulaCells(true);
            value = dataFormatter.formatCellValue(cell);
        }

        // 自定义的解析器
//        if (CellType.FORMULA.equals(cell.getCellType())) {
//            value = parserCellFormulaValue(cell);
//        } else {
//            value = parseCellBaseValue(cell, cell.getCellType());
//        }
        return parserdCellValueBuilder
                .value(value)
                .build();
    }
}
