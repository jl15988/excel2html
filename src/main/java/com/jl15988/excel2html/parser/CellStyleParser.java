package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.Excel2HtmlUtil;
import com.jl15988.excel2html.converter.ColorConverter;
import com.jl15988.excel2html.html.CssStyle;
import com.jl15988.excel2html.model.parser.ParserdStyle;
import com.jl15988.excel2html.model.parser.ParserdStyleResult;
import com.jl15988.excel2html.model.style.CommonCss;
import com.jl15988.excel2html.model.unit.UnitPixel;
import com.jl15988.excel2html.model.unit.UnitPoint;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

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
    public static ParserdStyle parserCellAlignGeneralStyle(CellType cellType, CellType formulaResultType) {
        ParserdStyle parserdStyle = new ParserdStyle();
        switch (cellType) {
            case NUMERIC:
                parserdStyle.styleMap.put("text-align", "right");
                parserdStyle.styleMap.put("justify-content", "flex-end");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_NUMERIC);
                break;
            case STRING:
                parserdStyle.styleMap.put("text-align", "left");
                parserdStyle.styleMap.put("justify-content", "flex-start");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_STRING);
                break;
            case BOOLEAN:
                parserdStyle.styleMap.put("text-align", "center");
                parserdStyle.styleMap.put("justify-content", "center");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_GENERAL_BOOLEAN);
                break;
        }
        if (Objects.nonNull(formulaResultType) && cellType == CellType.FORMULA) {
            ParserdStyle parserdStyleFormula = parserCellAlignGeneralStyle(formulaResultType, null);
            parserdStyle.merge(parserdStyleFormula);
        }
        return parserdStyle;
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

    /**
     * 解析单元格横向对齐方式样式
     *
     * @param cell 单元格
     * @return 样式
     */
    public static ParserdStyle parserCellHorizontalAlignStyle(Cell cell) {
        ParserdStyle parserdStyle = new ParserdStyle();
        // 水平对齐方式
        HorizontalAlignment alignment = cell.getCellStyle().getAlignment();
        switch (alignment) {
            case GENERAL:
                // 通用对齐方式。通常，文本数据左对齐，数字、日期和时间右对齐，布尔类型居中。
                CellType formulaResultType = null;
                if (cell.getCellType() == CellType.FORMULA) {
                    formulaResultType = cell.getCachedFormulaResultType();
                }
                ParserdStyle parserdStyleGeneral = parserCellAlignGeneralStyle(cell.getCellType(), formulaResultType);
                parserdStyle.merge(parserdStyleGeneral);
                break;
            case LEFT:
                // 左对齐。单元格内容靠左边缘对齐。
                parserdStyle.styleMap.put("text-align", "left");
                parserdStyle.styleMap.put("justify-content", "flex-start");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_LEFT);
                break;
            case CENTER:
                // 居中对齐。单元格内容在水平方向上居中对齐。
                parserdStyle.styleMap.put("text-align", "center");
                parserdStyle.styleMap.put("justify-content", "center");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_CENTER);
                break;
            case RIGHT:
                // 右对齐。单元格内容靠右边缘对齐。
                parserdStyle.styleMap.put("text-align", "right");
                parserdStyle.styleMap.put("justify-content", "flex-end");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_RIGHT);
                break;
            case FILL:
                // 填充对齐。单元格内容将填充整个单元格的宽度。
                // todo 由于css不兼容问题无法实现
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_FILL);
                break;
            case JUSTIFY:
                // 两端对齐。单元格内容将左右对齐，并在中间填充空格以达到两端对齐的效果。换行时第一行两端对齐
                parserdStyle.styleMap.put("text-align", "justify");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_JUSTIFY);
                break;
            case CENTER_SELECTION:
                // 跨多个单元格居中对齐。当选择多个单元格并设置此对齐方式时，内容将在所选单元格的区域内居中对齐。
                parserdStyle.styleMap.put("text-align", "center");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_CENTER_SELECTION);
                break;
            case DISTRIBUTED:
                // 分散对齐。单元格中的每一行文本中的每个单词将均匀分布在单元格的宽度内，左右边距对齐。所有行两端对齐
                parserdStyle.styleMap.put("text-align", "justify");
                parserdStyle.styleMap.put("text-align-last", "justify");
                parserdStyle.classList.add(CommonCss.ALIGN_HORIZONTAL_DISTRIBUTED);
                break;
        }
        return parserdStyle;
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

    /**
     * 解析单元格纵向对齐方式样式
     *
     * @param cell 单元格
     * @return 样式
     */
    public static ParserdStyle parserCellVerticalAlignStyle(Cell cell) {
        ParserdStyle parserdStyle = new ParserdStyle();
        // 垂直对齐方式
        VerticalAlignment verticalAlignment = cell.getCellStyle().getVerticalAlignment();
        switch (verticalAlignment) {
            case TOP:
                // 顶部对齐
                parserdStyle.styleMap.put("vertical-align", "baseline");
                parserdStyle.styleMap.put("align-items", "flex-start");
                parserdStyle.classList.add(CommonCss.ALIGN_VERTICAL_TOP);
                break;
            case CENTER:
                parserdStyle.styleMap.put("vertical-align", "middle");
                parserdStyle.styleMap.put("align-items", "center");
                parserdStyle.classList.add(CommonCss.ALIGN_VERTICAL_CENTER);
                break;
            case BOTTOM:
                parserdStyle.styleMap.put("vertical-align", "bottom");
                parserdStyle.styleMap.put("align-items", "flex-end");
                parserdStyle.classList.add(CommonCss.ALIGN_VERTICAL_BOTTOM);
                break;
            case JUSTIFY:
                parserdStyle.classList.add(CommonCss.ALIGN_VERTICAL_JUSTIFY);
                break;
            case DISTRIBUTED:
                parserdStyle.classList.add(CommonCss.ALIGN_VERTICAL_DISTRIBUTED);
                break;
        }
        return parserdStyle;
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

    /**
     * 解析单元格不同位置的边框样式
     *
     * @param cell     单元格
     * @param position 位置
     * @return 样式
     */
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
        String rgbaString = ColorConverter.xSSFColorToRGBAString(xSSFBorderColor);
        if (Objects.nonNull(rgbaString) && !rgbaString.isEmpty()) {
            borderColor = rgbaString;
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
                styleMap.put(borderStyleName, "1.5pt solid " + borderColor);
                break;
            case DOUBLE:
                styleMap.put(borderStyleName, "1.5pt double " + borderColor);
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

    /**
     * 解析单元格边框样式
     *
     * @param cell 单元格
     * @return 样式
     */
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

    /**
     * 解析字体样式
     *
     * @param font 字体
     * @return 样式
     */
    public static CssStyle parserFontStyle(Font font) {
        CssStyle cssStyle = new CssStyle();

        if (Objects.isNull(font)) return cssStyle;

        // 颜色
        if (font instanceof XSSFFont) {
            XSSFFont xssfFont = (XSSFFont) font;
            XSSFColor xssfColor = xssfFont.getXSSFColor();
            String fontRgba = ColorConverter.xSSFColorToRGBAString(xssfColor);
            cssStyle.setIfExists("color", fontRgba);
            // 加粗
            if (xssfFont.getBold()) {
                cssStyle.setIfExists("font-weight", "bold");
            }
            // 大小
            short fontHeightInPoints = xssfFont.getFontHeightInPoints();
            cssStyle.setIfExists("font-size", new UnitPoint(fontHeightInPoints).toString());
            // 斜体
            if (xssfFont.getItalic()) {
                cssStyle.setIfExists("font-style", "italic");
            }
            // 删除线
            if (xssfFont.getStrikeout()) {
                cssStyle.setIfExists("text-decoration", "line-through");
            }
            // 下划线
            if (xssfFont.getUnderline() != Font.U_NONE) {
                cssStyle.setIfExists("text-decoration", "underline");
            }
            String fontName = xssfFont.getFontName();
            cssStyle.setIfExists("font-family", fontName);
        }

        return cssStyle;
    }

    /**
     * 解析单元格字体样式
     *
     * @param cell 单元格
     * @return 样式
     */
    public static CssStyle parserCellFontStyle(Cell cell) {
        int fontIndex = cell.getCellStyle().getFontIndex();
        return CellStyleParser.parserFontStyle(cell.getRow().getSheet().getWorkbook().getFontAt(fontIndex));
    }

    /**
     * 解析单元格样式
     *
     * @param cell 单元格
     * @param dpi  屏幕 dpi
     * @return 样式
     */
    public static ParserdStyleResult parserCellStyle(Cell cell, int dpi) {
        ParserdStyleResult parserdStyleResult = new ParserdStyleResult();

        Row row = cell.getRow();
        XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();

        // 行高
        float heightInPoints = row.getHeightInPoints();
        String cellHeightC = new UnitPoint(heightInPoints - new UnitPixel(3, dpi).toPoint().getValue(), dpi).toString();
        String cellHeight = new UnitPoint(heightInPoints, dpi).toString();
        parserdStyleResult.cellContainerStyle.put("height", cellHeightC);
        parserdStyleResult.cellContainerStyle.put("max-height", cellHeightC);
        parserdStyleResult.cellContainerStyle.put("min-height", cellHeightC);
        parserdStyleResult.cellStyle.put("height", cellHeight);
        // 列宽
        double columnWidthInPixels = Excel2HtmlUtil.getColumnWidthInPixels(cell.getSheet(), cell.getColumnIndex());
        String cellWidth = new UnitPixel(columnWidthInPixels, dpi).toString();
        parserdStyleResult.cellStyle.put("width", cellWidth);
        parserdStyleResult.cellStyle.put("max-width", cellWidth);
        parserdStyleResult.cellStyle.put("min-width", cellWidth);

        // 对齐方式
        ParserdStyle horizontalAlignStyle = parserCellHorizontalAlignStyle(cell);
        parserdStyleResult.cellValCellStyle.putAll(horizontalAlignStyle.styleMap);
        parserdStyleResult.cellValStyleClassList.addAll(horizontalAlignStyle.classList);
        ParserdStyle verticalAlignStyle = parserCellVerticalAlignStyle(cell);
        parserdStyleResult.cellValCellStyle.putAll(verticalAlignStyle.styleMap);
        parserdStyleResult.cellValStyleClassList.addAll(verticalAlignStyle.classList);

        // 边框
        Map<String, Object> cellBorderStyle = parserCellBorderStyle(cell);
        parserdStyleResult.cellStyle.putAll(cellBorderStyle);
        // 背景
        XSSFColor fillBgColorColor = cellStyle.getFillBackgroundColorColor();
        parserdStyleResult.putIfExists(ParserdStyleResult::getCellStyle, "background-color", ColorConverter.xSSFColorToRGBAString(fillBgColorColor));
        XSSFColor fillForegroundColor = cellStyle.getFillForegroundColorColor();
        parserdStyleResult.putIfExists(ParserdStyleResult::getCellStyle, "background-color", ColorConverter.xSSFColorToRGBAString(fillForegroundColor));
        // 换行
        boolean wrapText = cellStyle.getWrapText();
        if (wrapText) {
            parserdStyleResult.cellContainerStyle.put("word-break", "break-word");
            parserdStyleResult.cellContainerStyle.put("white-space", "pre-wrap");
            parserdStyleResult.cellClassList.add("wrap-cell");
        } else {
            parserdStyleResult.cellContainerStyle.put("white-space", "pre");
        }
        // 字体
        CssStyle fontCssStyle = CellStyleParser.parserCellFontStyle(cell);
        parserdStyleResult.cellContainerStyle.putAll(fontCssStyle.getMap());

        return parserdStyleResult;
    }
}
