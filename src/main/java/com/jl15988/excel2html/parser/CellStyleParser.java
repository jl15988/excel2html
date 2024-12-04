package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.converter.ColorConverter;
import com.jl15988.excel2html.converter.UnitConverter;
import com.jl15988.excel2html.model.parser.ParserdStyle;
import com.jl15988.excel2html.model.style.CommonCss;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.*;

/**
 * 单元格样式解析器
 *
 * @author Jalon
 * @since 2024/12/1 20:24
 **/
public class CellStyleParser {

    /**
     * 解析单元格通用对齐方式为样式
     *
     * @param cellType          单元格类型
     * @param formulaResultType 当前单元格公式结果类型
     * @return 样式
     */
    public static Map<String, Object> parserCellAlignGeneralStyle(CellType cellType, CellType formulaResultType) {
        Map<String, Object> styleMap = new HashMap<>();
        switch (cellType) {
            case NUMERIC:
                styleMap.put("text-align", "right");
                styleMap.put("justify-content", "flex-end");
                break;
            case STRING:
                styleMap.put("text-align", "left");
                styleMap.put("justify-content", "flex-start");
                break;
            case BOOLEAN:
                styleMap.put("text-align", "center");
                styleMap.put("justify-content", "center");
                break;
        }
        if (Objects.nonNull(formulaResultType) && cellType == CellType.FORMULA) {
            Map<String, Object> styleMap2 = parserCellAlignGeneralStyle(formulaResultType, null);
            styleMap.putAll(styleMap2);
        }
        return styleMap;
    }

    /**
     * 解析单元格通用对齐方式为样式 class
     *
     * @param cellType          单元格类型
     * @param formulaResultType 当前单元格公式结果类型
     * @return 样式
     */
    public static List<String> parserCellAlignGeneralStyleClass(CellType cellType, CellType formulaResultType) {
        List<String> classList = new ArrayList<>();
        switch (cellType) {
            case NUMERIC:
                classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_NUMERIC);
                break;
            case STRING:
                classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_STRING);
                break;
            case BOOLEAN:
                classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_BOOLEAN);
                break;
        }
        if (Objects.nonNull(formulaResultType) && cellType == CellType.FORMULA) {
            List<String> styleClass2 = parserCellAlignGeneralStyleClass(formulaResultType, null);
            classList.addAll(styleClass2);
        }
        return classList;
    }

    public static Map<String, Object> parserCellHorizontalAlignStyle(Cell cell) {
        // 水平对齐方式
        HorizontalAlignment alignment = cell.getCellStyle().getAlignment();
        Map<String, Object> styleMap = new HashMap<>();
        switch (alignment) {
            case GENERAL:
                // 通用对齐方式。通常，文本数据左对齐，数字、日期和时间右对齐，布尔类型居中。
                CellType formulaResultType = null;
                if (cell.getCellType() == CellType.FORMULA) {
                    formulaResultType = cell.getCachedFormulaResultType();
                }
                Map<String, Object> generalStyle = parserCellAlignGeneralStyle(cell.getCellType(), formulaResultType);
                styleMap.putAll(generalStyle);
                break;
            case LEFT:
                // 左对齐。单元格内容靠左边缘对齐。
                styleMap.put("text-align", "left");
                styleMap.put("justify-content", "flex-start");
                break;
            case CENTER:
                // 居中对齐。单元格内容在水平方向上居中对齐。
                styleMap.put("text-align", "center");
                styleMap.put("justify-content", "center");
                break;
            case RIGHT:
                // 右对齐。单元格内容靠右边缘对齐。
                styleMap.put("text-align", "right");
                styleMap.put("justify-content", "flex-end");
                break;
            case FILL:
                // 填充对齐。单元格内容将填充整个单元格的宽度。
                // todo 由于css不兼容问题无法实现
                break;
            case JUSTIFY:
                // 两端对齐。单元格内容将左右对齐，并在中间填充空格以达到两端对齐的效果。换行时第一行两端对齐
                styleMap.put("text-align", "justify");
                break;
            case CENTER_SELECTION:
                // 跨多个单元格居中对齐。当选择多个单元格并设置此对齐方式时，内容将在所选单元格的区域内居中对齐。
                styleMap.put("text-align", "center");
                break;
            case DISTRIBUTED:
                // 分散对齐。单元格中的每一行文本中的每个单词将均匀分布在单元格的宽度内，左右边距对齐。所有行两端对齐
                styleMap.put("text-align", "justify");
                styleMap.put("text-align-last", "justify");
                break;
        }
        return styleMap;
    }

    public static List<String> parserCellHorizontalAlignStyleClass(Cell cell) {
        // 水平对齐方式
        HorizontalAlignment alignment = cell.getCellStyle().getAlignment();
        List<String> classList = new ArrayList<>();
        switch (alignment) {
            case GENERAL:
                // 通用对齐方式。通常，文本数据左对齐，数字、日期和时间右对齐，布尔类型居中。
                CellType formulaResultType = null;
                if (cell.getCellType() == CellType.FORMULA) {
                    formulaResultType = cell.getCachedFormulaResultType();
                }
                List<String> generalStyle = parserCellAlignGeneralStyleClass(cell.getCellType(), formulaResultType);
                classList.addAll(generalStyle);
                break;
            case LEFT:
                // 左对齐。单元格内容靠左边缘对齐。
                classList.add(CommonCss.ALIGN_HORIZONTAL_LEFT);
                break;
            case CENTER:
                // 居中对齐。单元格内容在水平方向上居中对齐。
                classList.add(CommonCss.ALIGN_HORIZONTAL_CENTER);
                break;
            case RIGHT:
                // 右对齐。单元格内容靠右边缘对齐。
                classList.add(CommonCss.ALIGN_HORIZONTAL_RIGHT);
                break;
            case FILL:
                // 填充对齐。单元格内容将填充整个单元格的宽度。
                // todo 由于css不兼容问题无法实现
                classList.add(CommonCss.ALIGN_HORIZONTAL_FILL);
                break;
            case JUSTIFY:
                // 两端对齐。单元格内容将左右对齐，并在中间填充空格以达到两端对齐的效果。换行时第一行两端对齐
                classList.add(CommonCss.ALIGN_HORIZONTAL_JUSTIFY);
                break;
            case CENTER_SELECTION:
                // 跨多个单元格居中对齐。当选择多个单元格并设置此对齐方式时，内容将在所选单元格的区域内居中对齐。
                classList.add(CommonCss.ALIGN_HORIZONTAL_CENTER_SELECTION);
                break;
            case DISTRIBUTED:
                // 分散对齐。单元格中的每一行文本中的每个单词将均匀分布在单元格的宽度内，左右边距对齐。所有行两端对齐
                classList.add(CommonCss.ALIGN_HORIZONTAL_DISTRIBUTED);
                break;
        }
        return classList;
    }

    public static Map<String, Object> parserCellVerticalAlignStyle(Cell cell) {
        Map<String, Object> styleMap = new HashMap<>();
        // 垂直对齐方式
        VerticalAlignment verticalAlignment = cell.getCellStyle().getVerticalAlignment();
        switch (verticalAlignment) {
            case TOP:
                // 顶部对齐
                styleMap.put("vertical-align", "baseline");
                styleMap.put("align-items", "flex-start");
                break;
            case CENTER:
                styleMap.put("vertical-align", "middle");
                styleMap.put("align-items", "center");
                break;
            case BOTTOM:
                styleMap.put("vertical-align", "bottom");
                styleMap.put("align-items", "flex-end");
                break;
            case JUSTIFY:
            case DISTRIBUTED:
        }
        return styleMap;
    }

    public static List<String> parserCellVerticalAlignStyleClass(Cell cell) {
        List<String> classList = new ArrayList<>();
        // 垂直对齐方式
        VerticalAlignment verticalAlignment = cell.getCellStyle().getVerticalAlignment();
        switch (verticalAlignment) {
            case TOP:
                // 顶部对齐
                classList.add(CommonCss.ALIGN_VERTICAL_TOP);
                break;
            case CENTER:
                classList.add(CommonCss.ALIGN_VERTICAL_CENTER);
                break;
            case BOTTOM:
                classList.add(CommonCss.ALIGN_VERTICAL_BOTTOM);
                break;
            case JUSTIFY:
                classList.add(CommonCss.ALIGN_VERTICAL_JUSTIFY);
                break;
            case DISTRIBUTED:
                classList.add(CommonCss.ALIGN_VERTICAL_DISTRIBUTED);
                break;
        }
        return classList;
    }

    public static Map<String, Object> parserCellBorderTypeStyle(Cell cell, String position) {
        Map<String, Object> styleMap = new HashMap<>();
        String borderStyleName = "border" + (position != null ? "-" + position : "");
        String borderColor = "black";
        XSSFColor xSSFBorderColor = null;
        XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
        BorderStyle borderStyle = null;
        if ("top".equals(position)) {
            borderStyle = cellStyle.getBorderTop();
            xSSFBorderColor = cellStyle.getTopBorderXSSFColor();
        } else if ("right".equals(position)) {
            borderStyle = cellStyle.getBorderRight();
            xSSFBorderColor = cellStyle.getRightBorderXSSFColor();
        } else if ("bottom".equals(position)) {
            borderStyle = cellStyle.getBorderBottom();
            xSSFBorderColor = cellStyle.getBottomBorderXSSFColor();
        } else if ("left".equals(position)) {
            borderStyle = cellStyle.getBorderLeft();
            xSSFBorderColor = cellStyle.getLeftBorderXSSFColor();
        }
        if (Objects.nonNull(xSSFBorderColor)) {
            String argbHex = xSSFBorderColor.getARGBHex();
            if (argbHex != null && !argbHex.isEmpty()) {
                borderColor = ColorConverter.argbHexToRgba(argbHex);
            }
        }

        if (Objects.isNull(borderStyle)) return styleMap;

        switch (borderStyle) {
            case NONE:
                styleMap.put(borderStyleName, "none");
                break;
            case THIN:
                styleMap.put(borderStyleName, "1px solid " + borderColor);
                break;
            case MEDIUM:
                styleMap.put(borderStyleName, "2px solid " + borderColor);
                break;
            case DASHED:
                styleMap.put(borderStyleName, "1px dashed " + borderColor);
                break;
            case DOTTED:
                styleMap.put(borderStyleName, "1px dotted " + borderColor);
                break;
            case THICK:
                styleMap.put(borderStyleName, "3px solid " + borderColor);
                break;
            case DOUBLE:
                styleMap.put(borderStyleName, "3px double " + borderColor);
                break;
            case HAIR:
                // todo 极细处理
                styleMap.put(borderStyleName, "1px solid " + borderColor);
                break;
            case MEDIUM_DASHED:
                styleMap.put(borderStyleName, "2px dashed " + borderColor);
                break;
            case DASH_DOT:
                styleMap.put(borderStyleName, "1px dash-dot " + borderColor);
                break;
            case MEDIUM_DASH_DOT:
                styleMap.put(borderStyleName, "2px dash-dot " + borderColor);
                break;
            case DASH_DOT_DOT:
                styleMap.put(borderStyleName, "1px dash-dot-dot " + borderColor);
                break;
            case MEDIUM_DASH_DOT_DOT:
                styleMap.put(borderStyleName, "2px dash-dot-dot " + borderColor);
                break;
            case SLANTED_DASH_DOT:
                styleMap.put(borderStyleName, "1px slanted dash-dot " + borderColor);
                break;
        }
        return styleMap;
    }

    public static Map<String, Object> parserCellBorderStyle(Cell cell) {
        Map<String, Object> styleMap = new HashMap<>();

        Map<String, Object> cellBorderTopStyle = parserCellBorderTypeStyle(cell, "top");
        styleMap.putAll(cellBorderTopStyle);
        Map<String, Object> cellBorderRightStyle = parserCellBorderTypeStyle(cell, "right");
        styleMap.putAll(cellBorderRightStyle);
        Map<String, Object> cellBorderBottomStyle = parserCellBorderTypeStyle(cell, "bottom");
        styleMap.putAll(cellBorderBottomStyle);
        Map<String, Object> cellBorderLeftStyle = parserCellBorderTypeStyle(cell, "left");
        styleMap.putAll(cellBorderLeftStyle);
        return styleMap;
    }

    public static ParserdStyle parserCellStyle(Cell cell, boolean compressStyle) {
        Map<String, Object> cellStyleMap = new HashMap<>();
        Map<String, Object> cellContainerStyleMap = new HashMap<>();
        Map<String, Object> cellValCellStyleMap = new HashMap<>();
        List<String> cellValClassList = new ArrayList<>();

        Row row = cell.getRow();
        Sheet sheet = row.getSheet();
        Workbook workbook = sheet.getWorkbook();

        XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();

        // 行高
        float heightInPoints = row.getHeightInPoints();
        String cellHeight = UnitConverter.convert().convertPointsString(heightInPoints);
        cellContainerStyleMap.put("height", cellHeight);
        cellContainerStyleMap.put("max-height", cellHeight);
        cellContainerStyleMap.put("min-height", cellHeight);
        // 列宽
        int columnIndex = cell.getColumnIndex();
        float columnWidthInPixels = sheet.getColumnWidthInPixels(columnIndex);
        String cellWidth = UnitConverter.convert().usePx().convertCellPixelsString(columnWidthInPixels);
        cellStyleMap.put("width", cellWidth);
        cellStyleMap.put("max-width", cellWidth);
        cellStyleMap.put("min-width", cellWidth);

        // 对齐方式
        if (compressStyle) {
            List<String> horizontalAlignStyle = parserCellHorizontalAlignStyleClass(cell);
            cellValClassList.addAll(horizontalAlignStyle);
            List<String> verticalAlignStyle = parserCellVerticalAlignStyleClass(cell);
            cellValClassList.addAll(verticalAlignStyle);
        } else {
            Map<String, Object> horizontalAlignStyle = parserCellHorizontalAlignStyle(cell);
            cellValCellStyleMap.putAll(horizontalAlignStyle);
            Map<String, Object> verticalAlignStyle = parserCellVerticalAlignStyle(cell);
            cellValCellStyleMap.putAll(verticalAlignStyle);
        }

        // 边框
        Map<String, Object> cellBorderStyle = parserCellBorderStyle(cell);
        cellStyleMap.putAll(cellBorderStyle);
        // 背景
//        XSSFColor fillForegroundColor = cellStyle.getFillForegroundColorColor();
//        if (Objects.nonNull(fillForegroundColor)) {
//            String argbHex = fillForegroundColor.getARGBHex();
//            if (argbHex != null && !argbHex.isEmpty()) {
//                cellStyleMap.put("background-color", argbHex);
//            }
//        }
//        XSSFColor fillBgColorColor = cellStyle.getFillBackgroundColorColor();
//        if (Objects.nonNull(fillBgColorColor)) {
//            String argbHex = fillBgColorColor.getARGBHex();
//            if (argbHex != null && !argbHex.isEmpty()) {
//                cellStyleMap.put("background-color", argbHex);
//            }
//        }
////        System.out.println(cellStyle.getFillBackgroundColor());
//        System.out.println(Arrays.toString(DefaultIndexedColorMap.getDefaultRGB(cellStyle.getFillForegroundColor())));
//        System.out.println(cellStyle.getFillBack);
        // 换行
        boolean wrapText = cellStyle.getWrapText();
        if (wrapText) {
            cellContainerStyleMap.put("word-break", "break-word");
            cellContainerStyleMap.put("white-space", "pre-wrap");
        } else {
            cellContainerStyleMap.put("white-space", "pre");
        }
        // 字体
        int fontIndex = cellStyle.getFontIndex();
        XSSFFont fontAt = (XSSFFont) workbook.getFontAt(fontIndex);
        // 加粗
        if (fontAt.getBold()) {
            cellContainerStyleMap.put("font-weight", "bold");
        }
        // 大小
        short fontHeightInPoints = fontAt.getFontHeightInPoints();
        cellContainerStyleMap.put("font-size", UnitConverter.convert().convertPointsString(fontHeightInPoints));
        // 斜体
        if (fontAt.getItalic()) {
            cellContainerStyleMap.put("font-style", "italic");
        }
        // 删除线
        if (fontAt.getStrikeout()) {
            cellContainerStyleMap.put("text-decoration", "line-through");
        }
        // 下划线
        if (fontAt.getUnderline() != Font.U_NONE) {
            cellContainerStyleMap.put("text-decoration", "underline");
        }
        String fontName = fontAt.getFontName();
        cellContainerStyleMap.put("font-family", fontName);

        return new ParserdStyle(cellStyleMap, cellContainerStyleMap, cellValCellStyleMap, cellValClassList);
    }
}
