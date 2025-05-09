<h1 align="center">excel2html</h1>

<p align="center">一个 excel 转 html 的 java 工具，旨在尽可能的还原 excel 原本的样式。<br>A Java tool for converting Excel to HTML, aimed at restoring the original style of Excel as much as possible.</p>

<p align="center">
	<a target="_blank" href="https://search.maven.org/artifact/com.jl15988.excel2html/excel2html">
		<img src="https://img.shields.io/maven-central/v/com.jl15988.excel2html/excel2html.svg?label=Maven%20Central" />
	</a>
    <img src="https://img.shields.io/:license-MIT-green.svg" />
</p>

## 使用

### 需要注意的是
- 目前仅支持 `xlsx` 格式
- 默认 dpi 是 96，因为不同屏幕的 dpi 可能不太一样（大多数是 96），所以尽量前端传过来


### 引入依赖
```xml
<dependency>
    <groupId>com.jl15988.excel2html</groupId>
    <artifactId>excel2html</artifactId>
    <version>0.0.1</version>
</dependency>
```

### 基本使用方法

#### 简单方式（使用辅助类）

```java
// 使用Excel2HtmlHelper快速将整个Sheet转换为HTML
HtmlPage htmlPage = Excel2HtmlHelper.toHtml(sheet);

// 转换为HTML字符串
String htmlString = htmlPage.toHtmlString();
```

#### 带参数方式（使用辅助类）

```java
// 指定行列范围和是否压缩样式
HtmlPage htmlPage = Excel2HtmlHelper.toHtml(sheet, 0, 20, 0, 10, true);

// 指定行列范围、是否压缩样式，并提供自定义的单元格格式化器和嵌入文件映射
HtmlPage htmlPage = Excel2HtmlHelper.toHtml(sheet, 0, 20, 0, 10, true, myCellValueFormatter, myEmbedFileMap);
```

#### 完整方式（使用核心类）

```java
// 创建Excel2Html实例（可从File、InputStream或byte[]创建）
Excel2Html excel2Html = new Excel2Html(new File("path/to/excel.xlsx"))
    // 设置屏幕DPI值（用于计算像素）
    .setDpi(96)
    // 设置是否压缩样式（将重复样式合并为CSS类）
    .setCompressStyle(true)
    // 设置是否加载嵌入文件（如图片）
    .setLoadEmbedFile(true)
    // 设置是否使用纸张模式（根据纸张大小限制行列）
    .setPaperMode(210f, 297f) // A4纸张尺寸（毫米）
    // 设置单元格值的格式化器
    .setCellValueFormater(new ICellValueFormater() {
        @Override
        public ParserdCellValue format(Cell cell, int rowIndex, int cellIndex) {
            // 自定义单元格值格式化逻辑
            return null; // 返回null表示使用默认格式化
        }
    })
    // 设置表格行元素处理器
    .setTrElementHandler(new ITrElementHandler() {
        @Override
        public void handle(HtmlElement tr, int rowIndex) {
            // 自定义处理表格行元素的逻辑
            tr.addClass("my-custom-row");
        }
    })
    // 设置单元格处理器
    .setCellHandler(new ICellHandler() {
        @Override
        public void handleStyle(ParserdStyleResult parserdStyleResult, Cell cell, int rowIndex, int cellIndex) {
            // 自定义处理单元格样式的逻辑
            if (rowIndex == 0) {
                // 为第一行添加特殊样式
                parserdStyleResult.cellStyle.put("background-color", "#f5f5f5");
            }
        }
    });

// 构建指定Sheet的HTML（可以指定行列范围）
HtmlPage htmlPage = excel2Html.buildHtml(sheet, 0, 20, 0, 10);

// 或者通过Sheet索引构建HTML
HtmlPage htmlPage = excel2Html.buildHtmlWithSheetIndex(0, 0, 20, 0, 10);

// 构建多个Sheet的HTML
List<HtmlPage> htmlPages = excel2Html.buildHtmlWithSheetIndex(0, 3, 0, 20, 0, 10);
```

#### 处理构建后的HTML

```java
// 获取HTML字符串（默认包含完整的HTML结构）
String htmlString = htmlPage.toHtmlString();

// 获取不包含HTML容器的字符串（仅包含表格内容）
String htmlContentString = htmlPage.setHasHtmlContainer(false).toHtmlString();

// 批量处理多个Sheet的HTML
List<String> htmlStrings = htmlPages.stream()
    .map(page -> page.setHasHtmlContainer(false).toHtmlString())
    .collect(Collectors.toList());
```

### 主要功能和配置选项

#### Excel2Html类的主要方法

- **setDpi(int dpi)**: 设置屏幕DPI值，用于计算像素转换，默认为96
- **setCompressStyle(boolean compressStyle)**: 设置是否启用样式压缩，默认为true
- **setLoadEmbedFile(boolean loadEmbedFile)**: 设置是否加载嵌入文件（如图片），默认为true
- **setPaperMode(Float paperWidth, Float paperHeight)**: 设置是否按纸张大小转换，指定纸张宽度和高度（单位：毫米）
- **setCellValueFormater(ICellValueFormater formater)**: 设置单元格值格式化器
- **setTrElementHandler(ITrElementHandler handler)**: 设置表格行元素处理器
- **setCellHandler(ICellHandler handler)**: 设置单元格处理器

#### 构建HTML的方法

- **buildHtml(Sheet sheet)**: 构建整个Sheet的HTML
- **buildHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)**: 构建指定行列范围的HTML
- **buildHtmlWithSheetIndex(int sheetIndex)**: 根据Sheet索引构建整个Sheet的HTML
- **buildHtmlWithSheetIndex(int sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)**: 根据Sheet索引构建指定行列范围的HTML
- **buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex)**: 构建多个Sheet的HTML
- **buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)**: 构建多个Sheet的指定行列范围的HTML

### 高级功能

#### 自定义单元格值格式化

通过实现`ICellValueFormater`接口可以自定义单元格值的格式化逻辑：

```java
excel2Html.setCellValueFormater(new ICellValueFormater() {
    @Override
    public ParserdCellValue format(Cell cell, int rowIndex, int cellIndex) {
        // 对特定单元格应用自定义格式
        if (rowIndex == 1 && cellIndex == 2) {
            ParserdCellValue value = new ParserdCellValue();
            value.text = "自定义内容";
            value.type = ParserdCellValueType.STRING;
            return value;
        }
        return null; // 返回null表示使用默认格式化
    }
});
```

#### 自定义单元格样式

通过实现`ICellHandler`接口可以自定义单元格样式：

```java
excel2Html.setCellHandler(new ICellHandler() {
    @Override
    public void handleStyle(ParserdStyleResult parserdStyleResult, Cell cell, int rowIndex, int cellIndex) {
        // 修改单元格样式
        if (rowIndex % 2 == 0) {
            // 偶数行背景色
            parserdStyleResult.cellStyle.put("background-color", "#f9f9f9");
        }
        
        // 去掉某些边框
        if (rowIndex == 4) {
            parserdStyleResult.cellStyle.remove("border-top");
        }
    }
});
```

#### 自定义行元素处理

通过实现`ITrElementHandler`接口可以自定义表格行元素：

```java
excel2Html.setTrElementHandler(new ITrElementHandler() {
    @Override
    public void handle(HtmlElement tr, int rowIndex) {
        // 为表格行添加自定义类或属性
        tr.addClass("row-" + rowIndex);
        
        if (rowIndex == 0) {
            tr.addClass("header-row");
        }
    }
});
```

## 难点

使用的是 `apache.poi` 依赖读取 excel，该依赖仍有某些不足，成为转 html 难点

1. 读取 excel 图片。excel 中图片有两种，第一种是浮动式，第二种是嵌入式，浮动式还好说 poi 能读取到，但是嵌入式只能自己解析 excel 内容，然后找到对应图片。excel 其实是一个压缩包，将其解压读取 xml 配置即可；
2. 渲染图片位置。因为获取到的浮动式图片位置为 emu 单位，且是所在单元格坐标的信息，单位转换和坐标计算有所难点；
3. 列宽。poi 读取到的列宽不准确，poi 中默认列宽写死了一个 8（字符宽度），这个 8 只是大概值，准确值需要自己计算；而且 poi 像素值都是乘了一个写死的 7.001699924468994（字符像素大小），这个值也是不准确的，这个值应该是 excel 默认字体的像素大小（一般国内都是默认宋体，像素大小为 8，差距也有点儿大），这个需要建立映射表，通过脚本将系统所有字体像素大小放到映射中，使用的时候再读取；
4. 富文本解析。富文本是指在同一个单元格使用不同的字体样式。这个需要对单元格内容单独解析，构造 html 样式，这个难点不算太大；
5. 空白字符处理。在 excel 中，连续空白字符是保留的，html 默认只显示一个，需要单独写样式，这个比较简单；如果单元格内容尾部含有空白字符且自动换行，空白字符是不占用空间的（目前看是这样），这个需要单独判断；
6. poi 获取单元格数据格式缺失，导致日期等格式格式化可能错误问题，通过补充数据格式增强格式化功能；
7. poi 的 DateUtil.isCellDateFormatted 方法不准确，增加了额外静态方法增强判断。
