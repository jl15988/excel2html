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

### 使用
```java
 List<HtmlPage> htmlPages = new Excel2Html(new File(respVO.getTempPath()))
        .setDpi(dpi)
        .setCellHandler(new ICellHandler() {
            @Override
            public void handleStyle(ParserdStyleResult parserdStyleResult, Cell cell, int rowIndex, int cellIndex) {
                // 去掉第一行单元格顶部边框
                if (rowIndex == 4) {
                    parserdStyleResult.cellStyle.remove("border-top");
                }
            }
        })
        .buildHtmlWithSheetIndex(4, null, 4, 46, 0, 29);
List<String> wbContent = htmlPages.stream().map(htmlPage -> htmlPage.setHasHtmlContainer(false).toHtmlString()).collect(Collectors.toList());
```

### 所有方法
- buildHtml(Sheet sheet, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)
- buildHtmlWithSheetIndex(int sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)
- buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColIndex, Integer endColIndex)
- buildHtml(Sheet sheet)
- buildHtmlWithSheetIndex(int sheetIndex)
- buildHtmlWithSheetIndex(Integer startSheetIndex, Integer endSheetIndex)

## 难点

使用的是 `apache.poi` 依赖读取 excel，该依赖仍有某些不足，成为转 html 难点

1. 读取 excel 图片。excel 中图片有两种，第一种是浮动式，第二种是嵌入式，浮动式还好说 poi 能读取到，但是嵌入式只能自己解析 excel 内容，然后找到对应图片。excel 其实是一个压缩包，将其解压读取 xml 配置即可；
2. 渲染图片位置。因为获取到的浮动式图片位置为 emu 单位，且是所在单元格坐标的信息，单位转换和坐标计算有所难点；
3. 列宽。poi 读取到的列宽不准确，poi 中默认列宽写死了一个 8（字符宽度），这个 8 只是大概值，准确值需要自己计算；而且 poi 像素值都是乘了一个写死的 7.001699924468994（字符像素大小），这个值也是不准确的，这个值应该是 excel 默认字体的像素大小（一般国内都是默认宋体，像素大小为 8，差距也有点儿大），这个需要建立映射表，通过脚本将系统所有字体像素大小放到映射中，使用的时候再读取；
4. 富文本解析。富文本是指在同一个单元格使用不同的字体样式。这个需要对单元格内容单独解析，构造 html 样式，这个难点不算太大；
5. 空白字符处理。在 excel 中，连续空白字符是保留的，html 默认只显示一个，需要单独写样式，这个比较简单；如果单元格内容尾部含有空白字符且自动换行，空白字符是不占用空间的（目前看是这样），这个需要单独判断。
