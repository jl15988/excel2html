package com.jl15988.excel2html.formatter;

import com.jl15988.excel2html.html.HtmlElement;
import org.apache.poi.ss.usermodel.Cell;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Jalon
 * @since 2024/12/5 12:55
 **/
public class DefaultCellValueFormatter implements ICellValueFormater {

    @Override
    public String format(String value, Cell cell) {
        Matcher matcher = Pattern.compile("\\s+$").matcher(value);
        if (matcher.find()) {
            String group = matcher.group();
            return matcher.replaceAll("") + HtmlElement.builder("span")
                    .addClass("value-end-spaces")
                    .content(group)
                    .build()
                    .toHtmlString();
        }
        return value;
    }
}
