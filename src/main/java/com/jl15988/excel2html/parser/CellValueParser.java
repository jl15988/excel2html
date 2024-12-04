package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.enums.ParserdCellValueType;
import com.jl15988.excel2html.formater.CellValueFormater;
import com.jl15988.excel2html.html.HtmlElement;
import com.jl15988.excel2html.html.HtmlElementList;
import com.jl15988.excel2html.model.parser.ParserdCellValue;
import com.jl15988.excel2html.model.style.FontICssStyle;
import org.apache.poi.ss.usermodel.*;
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
     */
    public static HtmlElementList parserCellRichValue(Cell cell) {
        if (!isRichValue(cell)) {
            return null;
        }
        HtmlElementList htmlElementList = new HtmlElementList();
        // 获取富文本内容
        XSSFRichTextString richText = (XSSFRichTextString) cell.getRichStringCellValue();
        int formattingRuns = richText.numFormattingRuns();
        for (int i = 0; i < formattingRuns; i++) {
            int start = richText.getIndexOfFormattingRun(i);
            int length = richText.getLengthOfFormattingRun(i);
            XSSFFont xssfFont = richText.getFontOfFormattingRun(i);
            String text = richText.toString().substring(start, start + length);

            // 解析字体样式
            FontICssStyle fontCssStyle = XSSFFontParser.parserXSSFFontToStyleMap(xssfFont);

            // 构建 span 标签
            HtmlElement spanEle = new HtmlElement("span");
            spanEle.addStyle(fontCssStyle).setContent(text);
            htmlElementList.add(spanEle);
        }
        return htmlElementList;
    }

    public static String parserCellNumericValue(Cell cell, CellType cellType) {
        if (cellType != CellType.NUMERIC) {
            return null;
        }
        if (DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue().toString();
        } else {
            return CellValueFormater.formatNumericValue(cell, cell.getNumericCellValue());
        }
    }

    /**
     * 解析公式结果
     *
     * @param cell 单元格
     * @return 结果
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
                    resultValue = CellValueFormater.formatNumericValue(cell, cellValue.getNumberValue());
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
     */
    public static ParserdCellValue parseCellValue(Cell cell, Map<String, XSSFPictureData> embedFileMap) {
        ParserdCellValue.ParserdCellValueBuilder parserdCellValueBuilder = ParserdCellValue
                .builder()
                .type(ParserdCellValueType.TEXT);

        CellType cellType = cell.getCellType();
        if (CellType.STRING.equals(cellType)) {
            // 对于 string 类型，需要解析富文本
            HtmlElementList htmlElementList = parserCellRichValue(cell);
            if (Objects.isNull(htmlElementList)) {
                return parserdCellValueBuilder.value(cell.getStringCellValue()).build();
            } else {
                return parserdCellValueBuilder
                        .type(ParserdCellValueType.RICH_HTML_CONTENT)
                        .value(htmlElementList.toHtmlString())
                        .build();
            }
        } else if (Objects.nonNull(embedFileMap) && CellType.FORMULA.equals(cellType)) {
            String cellFormula = cell.getCellFormula();
            if (cellFormula.startsWith("_xlfn.DISPIMG(\"")) {
                // 处理嵌入图片
                String embedValue = parserCellEmbedFile(cell, embedFileMap);
                if (Objects.nonNull(embedValue)) return parserdCellValueBuilder
                        .type(ParserdCellValueType.HTML_IMG)
                        .value(embedValue)
                        .build();
            }
        }
        // 官方提供的内容解析
        DataFormatter dataFormatter = new DataFormatter();
        dataFormatter.setUseCachedValuesForFormulaCells(true);
        return parserdCellValueBuilder
                .value(dataFormatter.formatCellValue(cell))
                .build();
        // 下面是自定义的内容解析
//        if (CellType.FORMULA.equals(cell.getCellType())) {
//            return parserCellFormulaValue(cell);
//        } else {
//            return parseCellBaseValue(cell, cell.getCellType());
//        }
    }
}