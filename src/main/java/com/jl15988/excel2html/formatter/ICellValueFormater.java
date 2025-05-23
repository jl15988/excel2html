package com.jl15988.excel2html.formatter;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 单元格内容格式化
 *
 * @author Jalon
 * @since 2024/12/5 11:26
 **/
public interface ICellValueFormater {

    String format(String value, Cell cell);
}
