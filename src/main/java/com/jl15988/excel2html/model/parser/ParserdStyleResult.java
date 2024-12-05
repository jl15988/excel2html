package com.jl15988.excel2html.model.parser;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Jalon
 * @since 2024/12/1 20:35
 **/
@Data
@AllArgsConstructor
public class ParserdStyleResult {

    public Map<String, Object> cellStyle;
    public Map<String, Object> cellContainerStyle;
    public Map<String, Object> cellValCellStyle;
    public List<String> cellValClassList;

    public ParserdStyleResult() {
        this.cellStyle = new HashMap<>();
        this.cellContainerStyle = new HashMap<>();
        this.cellValCellStyle = new HashMap<>();
        this.cellValClassList = new ArrayList<>();
    }

    public ParserdStyleResult addCellStyle(Map<String, Object> cellStyle) {
        if (this.cellStyle == null) {
            this.cellStyle = new HashMap<>();
        }
        this.cellStyle.putAll(cellStyle);
        return this;
    }

    public ParserdStyleResult addCellStyle(String cellStyleName, Object cellStyleValue) {
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

    public ParserdStyleResult addCellContainerStyle(Map<String, Object> cellContainerStyle) {
        if (this.cellContainerStyle == null) {
            this.cellContainerStyle = new HashMap<>();
        }
        this.cellContainerStyle.putAll(cellContainerStyle);
        return this;
    }

    public ParserdStyleResult addCellContainerStyle(String cellContainerStyleName, Object cellContainerStyleValue) {
        if (this.cellContainerStyle == null) {
            this.cellContainerStyle = new HashMap<>();
        }
        this.cellContainerStyle.put(cellContainerStyleName, cellContainerStyleValue);
        return this;
    }

    public ParserdStyleResult addCellValCellStyle(Map<String, Object> cellValCellStyle) {
        if (this.cellValCellStyle == null) {
            this.cellValCellStyle = new HashMap<>();
        }
        this.cellValCellStyle.putAll(cellValCellStyle);
        return this;
    }

    public ParserdStyleResult addCellValCellStyle(String cellValCellStyleName, String cellValCellStyleValue) {
        if (this.cellValCellStyle == null) {
            this.cellValCellStyle = new HashMap<>();
        }
        this.cellValCellStyle.put(cellValCellStyleName, cellValCellStyleValue);
        return this;
    }
}
