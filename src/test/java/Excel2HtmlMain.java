import com.jl15988.excel2html.Excel2Html;
import com.jl15988.excel2html.Excel2HtmlHelper;
import com.jl15988.excel2html.utils.FileUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
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

        String fileName = "测试记录表.xlsx";
        String resultName = fileName.substring(0, fileName.lastIndexOf(".")) + ".html";

        String userDir = System.getProperty("user.dir");
        String testFileDir = userDir + "\\src\\test\\java\\resources\\";

        String excelFilePath = testFileDir + fileName;
        String htmlFilePath = testFileDir + resultName;

        try {
            FileInputStream fis = new FileInputStream(excelFilePath);

            byte[] fileData = FileUtil.getFileStream(fis);
            assert fileData != null;
            FileOutputStream fos = new FileOutputStream(htmlFilePath);

//            XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(fileData));
//            String htmlString = Excel2HtmlHelper.toHtml(workbook.getSheetAt(0)).toHtmlString();

            // 如果你想解析 excel 中的图片，需要通过 excel 文件构建 Excel2Html 对象，或者通过 Excel2HtmlUtil 自己获取嵌入图片，然后在 Excel2HtmlHelper.toHtml 方法中传入
            Excel2Html excel2Html = new Excel2Html(new File(excelFilePath));
            String htmlString = excel2Html.buildHtmlWithSheetIndex(0).toHtmlString();

            fos.write((htmlString == null ? "" : htmlString).getBytes(StandardCharsets.UTF_8));

            System.out.println("转换完成，耗时：" + (System.currentTimeMillis() - startTime) / 1000f + "s");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
