package com.jl15988.excel2html.model.parser;

import com.jl15988.excel2html.enums.ParserdCellValueType;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author Jalon
 * @since 2024/12/3 11:35
 **/
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ParserdCellValue {

    private String value;

    private ParserdCellValueType type;
}
