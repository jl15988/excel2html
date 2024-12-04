package com.jl15988.excel2html.parser;

import com.jl15988.excel2html.converter.UnitConverter;
import com.jl15988.excel2html.html.HtmlElement;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

/**
 * 图形解析器
 *
 * @author Jalon
 * @since 2024/12/2 17:27
 **/
public class DrawingValueParser {

    /**
     * 解析表格中的形状
     * <p>
     * 因为只能获取到形状锚点位置，所以只能获取大体位置
     *
     * @param sheet 表格 sheet
     */
    public static List<HtmlElement> parserDrawing(Sheet sheet) {
        List<HtmlElement> htmlElementList = new ArrayList<>();
        // 获取形状，包含图片
        Drawing<?> drawingPatriarch = sheet.getDrawingPatriarch();
        if (drawingPatriarch != null) {
            for (Shape patriarch : drawingPatriarch) {
                if (patriarch instanceof XSSFPicture) {
                    // 图片类型
                    XSSFPicture pictureShape = (XSSFPicture) patriarch;

                    ClientAnchor anchor = pictureShape.getPreferredSize();
                    // 获取锚点所在单元格
                    int col1 = anchor.getCol1();
                    int row1 = anchor.getRow1();
                    int col2 = anchor.getCol2();
                    int row2 = anchor.getRow2();

                    UnitConverter dxUnitConverter = UnitConverter.convert().usePx();
                    UnitConverter dyUnitConverter = UnitConverter.convert().usePx();

                    // 获取锚点在所在单元格的坐标
                    double dx1 = dxUnitConverter.convertEmus(anchor.getDx1());
                    double dy1 = dyUnitConverter.convertEmus(anchor.getDy1());
                    double dx2 = dxUnitConverter.convertEmus(anchor.getDx2());
                    double dy2 = dyUnitConverter.convertEmus(anchor.getDy2());

                    byte[] imageBytes = pictureShape.getPictureData().getData();
                    String base64Image = "data:image/png;base64," + Base64.getEncoder().encodeToString(imageBytes);

                    HtmlElement img = new HtmlElement("img");
                    img.addAttribute("src", base64Image);
                    img.addStyle("width", totalColumnWidth(col1, col2, sheet, dxUnitConverter) + (dx2 - dx1) + dxUnitConverter.getUnit());
                    img.addStyle("height", totalRowHeight(row1, row2, sheet, dyUnitConverter) + (dy2 - dy1) + dyUnitConverter.getUnit());
                    img.addStyle("top", totalRowHeight(0, row1, sheet, dyUnitConverter) + dy1 + dyUnitConverter.getUnit());
                    img.addStyle("left", totalColumnWidth(0, col1, sheet, dxUnitConverter) + dx1 + dxUnitConverter.getUnit());
                    img.addStyle("position", "absolute");
                    htmlElementList.add(img);
                } else if (patriarch instanceof XSSFSimpleShape) {
                    // 其他形状
                }
            }
        }
        return htmlElementList;
    }

    public static double totalColumnWidth(int col1, int col2, Sheet sheet, UnitConverter dxUnitConverter) {
        double totalWidth = 0;
        for (int i = col1; i < col2; i++) {
            totalWidth += UnitConverter.convert().setUsePx(dxUnitConverter.getUsePx()).convertCellPixels(sheet.getColumnWidthInPixels(i));
        }
        return totalWidth;
    }

    public static double totalRowHeight(int row1, int row2, Sheet sheet, UnitConverter dyUnitConverter) {
        double totalHeight = 0;
        for (int i = row1; i < row2; i++) {
            totalHeight += UnitConverter.convert().setUsePx(dyUnitConverter.getUsePx()).convertPoints(sheet.getRow(i).getHeightInPoints());
        }
        return totalHeight;
    }

    private static void drawShape(XSSFSimpleShape simpleShape) {
        // 创建一个画布
        BufferedImage image = new BufferedImage(200, 200, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = image.createGraphics();

        // 设置背景颜色
        graphics.setPaint(Color.BLUE);
        XSSFAnchor anchor = simpleShape.getAnchor();
        int x = Units.pointsToPixel(Units.toPoints(anchor.getDx1()));
        int y = Units.pointsToPixel(Units.toPoints(anchor.getDy1()));
        int x2 = Units.pointsToPixel(Units.toPoints(anchor.getDx2()));
        int y2 = Units.pointsToPixel(Units.toPoints(anchor.getDy2()));

        System.out.println(x);
        System.out.println(y);
        System.out.println(x2);
        System.out.println(y2);

        graphics.fillRect(x, y, 20, y2 - y);
    }

    private static String bufferedImageToBase64(BufferedImage image) {
        ByteArrayOutputStream bao = new ByteArrayOutputStream();//io流
        try {
            //写入流中
            ImageIO.write(image, "png", bao);
        } catch (IOException e) {
            e.printStackTrace();
        }
        byte[] bytes = Base64.getEncoder().encode(bao.toByteArray());
        String base64 = new String(bytes);
        base64 = base64.replaceAll("\n", "").replaceAll("\r", "");//删除 \r\n
        String base64Image = "data:image/png;base64," + base64;
        return base64Image;
    }
}