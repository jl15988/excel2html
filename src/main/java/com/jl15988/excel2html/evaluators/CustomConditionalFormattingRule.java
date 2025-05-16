package com.jl15988.excel2html.evaluators;

import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionFilterData;
import org.apache.poi.ss.usermodel.ConditionFilterType;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataBarFormatting;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.poi.ss.usermodel.PatternFormatting;

import java.text.NumberFormat;

/**
 * @author Jalon
 * @since 2025/5/16 10:54
 **/
public class CustomConditionalFormattingRule implements ConditionalFormattingRule {

    private ExcelNumberFormat numberFormat;

    public void setNumberFormat(ExcelNumberFormat format) {
        this.numberFormat = format;
    }

    @Override
    public BorderFormatting createBorderFormatting() {
        return null;
    }

    @Override
    public BorderFormatting getBorderFormatting() {
        return null;
    }

    @Override
    public FontFormatting createFontFormatting() {
        return null;
    }

    @Override
    public FontFormatting getFontFormatting() {
        return null;
    }

    @Override
    public PatternFormatting createPatternFormatting() {
        return null;
    }

    @Override
    public PatternFormatting getPatternFormatting() {
        return null;
    }

    @Override
    public int getStripeSize() {
        return 0;
    }

    @Override
    public DataBarFormatting getDataBarFormatting() {
        return null;
    }

    @Override
    public IconMultiStateFormatting getMultiStateFormatting() {
        return null;
    }

    @Override
    public ColorScaleFormatting getColorScaleFormatting() {
        return null;
    }

    @Override
    public ExcelNumberFormat getNumberFormat() {
        return numberFormat;
    }

    @Override
    public ConditionType getConditionType() {
        return null;
    }

    @Override
    public ConditionFilterType getConditionFilterType() {
        return null;
    }

    @Override
    public ConditionFilterData getFilterConfiguration() {
        return null;
    }

    @Override
    public byte getComparisonOperation() {
        return 0;
    }

    @Override
    public String getFormula1() {
        return "";
    }

    @Override
    public String getFormula2() {
        return "";
    }

    @Override
    public String getText() {
        return "";
    }

    @Override
    public int getPriority() {
        return 0;
    }

    @Override
    public boolean getStopIfTrue() {
        return false;
    }
}
