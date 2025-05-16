package com.jl15988.excel2html.evaluators;

import com.jl15988.excel2html.parser.CellDataFormatParser;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;

import java.util.Collections;
import java.util.List;

/**
 * @author Jalon
 * @since 2025/5/16 10:42
 **/
public class CustomConditionalFormattingEvaluator extends ConditionalFormattingEvaluator {
    public CustomConditionalFormattingEvaluator(Workbook wb, WorkbookEvaluatorProvider provider) {
        super(wb, (WorkbookEvaluatorProvider) wb.getCreationHelper().createFormulaEvaluator());
    }

    @Override
    public List<EvaluationConditionalFormatRule> getConditionalFormattingForCell(Cell cell) {
        CustomConditionalFormattingRule customConditionalFormattingRule = new CustomConditionalFormattingRule();
        String dataFormatString = CellDataFormatParser.getDataFormatString(cell);
        customConditionalFormattingRule.setNumberFormat(new ExcelNumberFormat(cell.getCellStyle().getDataFormat(), dataFormatString));
        EvaluationConditionalFormatRule evaluationConditionalFormatRule = new EvaluationConditionalFormatRule(null, null, null, 0, customConditionalFormattingRule, 0, new CellRangeAddress[]{});
        return Collections.singletonList(evaluationConditionalFormatRule);
    }
}
