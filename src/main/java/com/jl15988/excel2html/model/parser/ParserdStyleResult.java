package com.jl15988.excel2html.model.parser;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.*;
import java.util.function.Function;

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
    // 单元格内容样式 class，与 class不同的是，样式 class 是按判断是否压缩来赋值的
    public List<String> cellValStyleClassList;

    // 单元格 class
    public List<String> cellClassList;

    public ParserdStyleResult() {
        this.cellStyle = new HashMap<>();
        this.cellContainerStyle = new HashMap<>();
        this.cellValCellStyle = new HashMap<>();
        this.cellValStyleClassList = new ArrayList<>();

        this.cellClassList = new ArrayList<>();
    }

    public ParserdStyleResult putIfExists(Function<ParserdStyleResult, Map<String, Object>> mapper, String name, Object value) {
        Map<String, Object> map = mapper.apply(this);
        if (Objects.nonNull(map) && Objects.nonNull(value)) {
            map.put(name, value);
        }
        return this;
    }

    public ParserdStyleResult addIfExists(Function<ParserdStyleResult, List<String>> mapper, String value) {
        List<String> list = mapper.apply(this);
        if (Objects.nonNull(list) && Objects.nonNull(value)) {
            list.add(value);
        }
        return this;
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
