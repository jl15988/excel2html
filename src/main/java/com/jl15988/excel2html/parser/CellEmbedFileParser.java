package com.jl15988.excel2html.parser;

import cn.hutool.core.util.StrUtil;
import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;
import cn.hutool.json.XML;
import org.apache.commons.collections4.map.HashedMap;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * 单元格嵌入文件解析器
 *
 * @author Jalon
 * @since 2024/12/3 10:36
 **/
public class CellEmbedFileParser {

    /**
     * 处理 WPS 文件中的图片数据，获取图片id-data
     *
     * @param stream    输入流
     * @param mapConfig 配置映射
     * @return 图片信息的 map
     * @throws IOException
     */
    public static Map<String, XSSFPictureData> processPictures(ByteArrayInputStream stream, Map<String, String> mapConfig) throws IOException {
        Map<String, XSSFPictureData> mapPictures = new HashedMap<>();
        Workbook workbook = WorkbookFactory.create(stream);
        List<XSSFPictureData> allPictures = (List<XSSFPictureData>) workbook.getAllPictures();
        for (XSSFPictureData pictureData : allPictures) {
            PackagePartName partName = pictureData.getPackagePart().getPartName();
            String uri = partName.getURI().toString();
            if (mapConfig.containsKey(uri)) {
                String strId = mapConfig.get(uri);
                mapPictures.put(strId, pictureData);
            }
        }
        return mapPictures;
    }

    /**
     * 处理 Zip 文件中的条目，获取图片路径-id
     *
     * @param stream Zip 输入流
     * @return 配置信息的 map
     * @throws IOException
     */
    public static Map<String, String> processZipEntries(ByteArrayInputStream stream) throws IOException {
        Map<String, String> mapConfig = new HashedMap<>();
        ZipInputStream zipInputStream = new ZipInputStream(stream);
        ZipEntry zipEntry;
        while ((zipEntry = zipInputStream.getNextEntry()) != null) {
            try {
                final String fileName = zipEntry.getName();
                if ("xl/cellimages.xml".equals(fileName)) {
                    processCellImages(zipInputStream, mapConfig);
                } else if ("xl/_rels/cellimages.xml.rels".equals(fileName)) {
                    return processCellImagesRels(zipInputStream, mapConfig);
                }
            } finally {
                zipInputStream.closeEntry();
            }
        }
        return new HashedMap<>();
    }

    /**
     * 处理 Zip 文件中的 cellimages.xml 文件，更新图片配置信息。
     *
     * @param zipInputStream Zip 输入流
     * @param mapConfig      配置信息的 map
     * @throws IOException
     */
    private static void processCellImages(ZipInputStream zipInputStream, Map<String, String> mapConfig) throws IOException {
        String content = IOUtils.toString(zipInputStream, StandardCharsets.UTF_8);
        JSONObject jsonObject = XML.toJSONObject(content);
        if (jsonObject != null) {
            JSONObject cellImages = jsonObject.getJSONObject("etc:cellImages");
            if (cellImages != null) {
                JSONArray cellImageArray = null;
                Object cellImage = cellImages.get("etc:cellImage");
                if (cellImage instanceof JSONArray) {
                    cellImageArray = (JSONArray) cellImage;
                } else if (cellImage instanceof JSONObject) {
                    JSONObject cellImageObj = (JSONObject) cellImage;
                    cellImageArray = new JSONArray();
                    cellImageArray.add(cellImageObj);
                }
                if (cellImageArray != null) {
                    processImageItems(cellImageArray, mapConfig);
                }
            }
        }
    }

    /**
     * 处理 cellImageArray 中的图片项，更新图片配置信息。
     *
     * @param cellImageArray 图片项的 JSONArray
     * @param mapConfig      配置信息的 map
     */
    private static void processImageItems(JSONArray cellImageArray, Map<String, String> mapConfig) {
        for (int i = 0; i < cellImageArray.size(); i++) {
            JSONObject imageItem = cellImageArray.getJSONObject(i);
            if (imageItem != null) {
                JSONObject pic = imageItem.getJSONObject("xdr:pic");
                if (pic != null) {
                    processPic(pic, mapConfig);
                }
            }
        }
    }

    /**
     * 处理 pic 中的图片信息，更新图片配置信息。
     *
     * @param pic       图片的 JSONObject
     * @param mapConfig 配置信息的 map
     */
    private static void processPic(JSONObject pic, Map<String, String> mapConfig) {
        JSONObject nvPicPr = pic.getJSONObject("xdr:nvPicPr");
        if (nvPicPr != null) {
            JSONObject cNvPr = nvPicPr.getJSONObject("xdr:cNvPr");
            if (cNvPr != null) {
                String name = cNvPr.getStr("name");
                if (StrUtil.isNotEmpty(name)) {
                    String strImageEmbed = updateImageEmbed(pic);
                    if (strImageEmbed != null) {
                        mapConfig.put(strImageEmbed, name);
                    }
                }
            }
        }
    }

    /**
     * 获取嵌入式图片的 embed 信息。
     *
     * @param pic 图片的 JSONObject
     * @return embed 信息
     */
    private static String updateImageEmbed(JSONObject pic) {
        JSONObject blipFill = pic.getJSONObject("xdr:blipFill");
        if (blipFill != null) {
            JSONObject blip = blipFill.getJSONObject("a:blip");
            if (blip != null) {
                return blip.getStr("r:embed");
            }
        }
        return null;
    }

    /**
     * 处理 Zip 文件中的 relationship 条目，更新配置信息。
     *
     * @param zipInputStream Zip 输入流
     * @param mapConfig      配置信息的 map
     * @return 配置信息的 map
     * @throws IOException
     */
    private static Map<String, String> processCellImagesRels(ZipInputStream zipInputStream, Map<String, String> mapConfig) throws IOException {
        String content = IOUtils.toString(zipInputStream, StandardCharsets.UTF_8);
        JSONObject jsonObject = XML.toJSONObject(content);
        JSONObject relationships = jsonObject.getJSONObject("Relationships");
        if (relationships != null) {
            JSONArray relationshipArray = null;
            Object relationship = relationships.get("Relationship");

            if (relationship instanceof JSONArray) {
                relationshipArray = (JSONArray) relationship;
            } else if (relationship instanceof JSONObject) {
                JSONObject relationshipObj = (JSONObject) relationship;
                relationshipArray = new JSONArray();
                relationshipArray.add(relationshipObj);
            }
            if (relationshipArray != null) {
                return processRelationships(relationshipArray, mapConfig);
            }
        }
        return null;
    }

    /**
     * 处理 relationshipArray 中的关系项，更新配置信息。
     *
     * @param relationshipArray 关系项的 JSONArray
     * @param mapConfig         配置信息的 map
     * @return 配置信息的 map
     */
    private static Map<String, String> processRelationships(JSONArray relationshipArray, Map<String, String> mapConfig) {
        Map<String, String> mapRelationships = new HashedMap<>();
        for (int i = 0; i < relationshipArray.size(); i++) {
            JSONObject relaItem = relationshipArray.getJSONObject(i);
            if (relaItem != null) {
                String id = relaItem.getStr("Id");
                String value = "/xl/" + relaItem.getStr("Target");
                if (mapConfig.containsKey(id)) {
                    String strImageId = mapConfig.get(id);
                    mapRelationships.put(value, strImageId);
                }
            }
        }
        return mapRelationships;
    }
}
