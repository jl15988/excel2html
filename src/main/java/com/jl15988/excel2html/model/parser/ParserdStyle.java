package com.jl15988.excel2html.model.parser;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/1 20:35
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ParserdStyle {

    private Map<String, Object> cellStyle;
    private Map<String, Object> cellContainerStyle;
    private Map<String, Object> cellValCellStyle;
    private List<String> cellValClassList;

    public ParserdStyle addCellStyle(Map<String, Object> cellStyle) {
        if (this.cellStyle == null) {
            this.cellStyle = new HashMap<>();
        }
        this.cellStyle.putAll(cellStyle);
        return this;
    }

    public ParserdStyle addCellStyle(String cellStyleName, Object cellStyleValue) {
        if (this.cellStyle == null) {
            this.cellStyle = new HashMap<>();
        }
        this.cellStyle.put(cellStyleName, cellStyleValue);
        return this;
    }

    public boolean hasCellStyle(String cellStyleName) {
        if (this.cellStyle == null) {
            return false;
        }
        return this.cellStyle.containsKey(cellStyleName);
    }

    public ParserdStyle addCellContainerStyle(Map<String, Object> cellContainerStyle) {
        if (this.cellContainerStyle == null) {
            this.cellContainerStyle = new HashMap<>();
        }
        this.cellContainerStyle.putAll(cellContainerStyle);
        return this;
    }

    public ParserdStyle addCellContainerStyle(String cellContainerStyleName, Object cellContainerStyleValue) {
        if (this.cellContainerStyle == null) {
            this.cellContainerStyle = new HashMap<>();
        }
        this.cellContainerStyle.put(cellContainerStyleName, cellContainerStyleValue);
        return this;
    }

    public ParserdStyle addCellValCellStyle(Map<String, Object> cellValCellStyle) {
        if (this.cellValCellStyle == null) {
            this.cellValCellStyle = new HashMap<>();
        }
        this.cellValCellStyle.putAll(cellValCellStyle);
        return this;
    }

    public ParserdStyle addCellValCellStyle(String cellValCellStyleName, String cellValCellStyleValue) {
        if (this.cellValCellStyle == null) {
            this.cellValCellStyle = new HashMap<>();
        }
        this.cellValCellStyle.put(cellValCellStyleName, cellValCellStyleValue);
        return this;
    }
}
