package com.jl15988.excel2html.handler;

import com.jl15988.excel2html.html.HtmlElement;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author Jalon
 * @since 2024/12/9 10:46
 **/
public interface ITrElementHandler {

    HtmlElement handle(HtmlElement tr, int rowIndex, Sheet sheet);
}
