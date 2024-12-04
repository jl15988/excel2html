package com.jl15988.excel2html.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;

/**
 * @author Jalon
 * @since 2024/12/3 10:53
 **/
public class FileUtil {

    /**
     * 获取文件数据
     *
     * @param inputStream 文件
     */
    public static byte[] getFileStream(InputStream inputStream) {
        try {
            // 创建 ByteArrayOutputStream 来暂存流数据
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            // 将 inputStream 读取到 byteArrayOutputStream 中
            byte[] buffer = new byte[1024];
            int length;
            while ((length = inputStream.read(buffer)) != -1) {
                byteArrayOutputStream.write(buffer, 0, length);
            }
            // 将 byteArrayOutputStream 的内容获取为字节数组
            return byteArrayOutputStream.toByteArray();
        } catch (IOException e) {
            return null;
        }
    }

    /**
     * 获取文件数据
     *
     * @param file 文件
     */
    public static byte[] getFileStream(File file) throws IOException {
        InputStream inputStream = Files.newInputStream(file.toPath());
        return getFileStream(inputStream);
    }
}
