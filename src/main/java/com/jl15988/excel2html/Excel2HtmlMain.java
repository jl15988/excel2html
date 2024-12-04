package com.jl15988.excel2html;

import com.jl15988.excel2html.utils.FileUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;

/**
 * @author Jalon
 * @since 2024/12/1 20:58
 **/
public class Excel2HtmlMain {

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();
        String excelFilePath = "D:\\developerSpace\\GitProjects\\excel2html\\src\\main\\resources\\测试记录表.xlsx";
        String htmlFilePath = "D:\\developerSpace\\GitProjects\\excel2html\\src\\main\\resources\\test.html";

        try {
            FileInputStream fis = new FileInputStream(excelFilePath);

            byte[] fileData = FileUtil.getFileStream(fis);
            assert fileData != null;
            FileOutputStream fos = new FileOutputStream(htmlFilePath);

            Excel2HtmlBuilder excel2HtmlBuilder = new Excel2HtmlBuilder(new File(excelFilePath));
            String htmlString = excel2HtmlBuilder.buildHtml(0).toHtmlString();

            fos.write((htmlString == null ? "" : htmlString).getBytes(StandardCharsets.UTF_8));

            System.out.println("转换完成，耗时：" + (System.currentTimeMillis() - startTime) / 1000f + "s");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
