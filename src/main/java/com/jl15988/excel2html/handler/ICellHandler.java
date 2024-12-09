package com.jl15988.excel2html.handler;

import com.jl15988.excel2html.html.HtmlElement;
import com.jl15988.excel2html.model.parser.ParserdStyleResult;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Jalon
 * @since 2024/12/9 10:46
 **/
public interface ICellHandler {

    default HtmlElement handle(HtmlElement td, int rowIndex, int cellIndex, Sheet sheet) {
        return td;
    }

    default void handleStyle(ParserdStyleResult parserdStyleResult, Cell cell, int rowIndex, int cellIndex) {
    }
}
