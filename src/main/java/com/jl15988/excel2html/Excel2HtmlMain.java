package com.jl15988.excel2html;

import com.jl15988.excel2html.utils.FileUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;

/**
 * @author Jalon
 * @since 2024/12/1 20:58
 **/
public class Excel2HtmlMain {

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        // 获取当前类加载器
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        // 获取resources目录的URL地址
        URL resource = classLoader.getResource("");
        assert resource != null;
        String resourcePath = resource.getPath().replace("target/classes/", "") + "src\\main\\resources\\";

        String excelFilePath = resourcePath + "测试记录表.xlsx";
        String htmlFilePath = resourcePath + "test.html";

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
