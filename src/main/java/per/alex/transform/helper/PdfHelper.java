package per.alex.transform.helper;


import com.lowagie.text.pdf.PdfReader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PdfHelper {

    private static final Logger logger = LoggerFactory.getLogger(PdfHelper.class);

    /**
     * 最大dpi限制
     */
    private static final Integer MAX_DPI = 300;

    /**
     * 默认dpi
     */
    private static final Integer DEFAULT_DPI = 240;

    /***
     * PDF文件转PNG图片，全部页数
     *
     * @param PdfFilePath pdf完整路径
     * @param dstImgFolder 图片存放的文件夹
     * @param dpi dpi越大转换后越清晰，相对转换速度越慢
     * @return
     */
    public static List<String>  convertImage(String PdfFilePath, String dstImgFolder, int dpi) {
        File file = new File(PdfFilePath);
        List<String> imageList = new ArrayList<>();
        PDDocument pdDocument;
        try {
            String imgPDFPath = file.getParent();
            int dot = file.getName().lastIndexOf('.');
            String imagePDFName = file.getName().substring(0, dot); // 获取图片文件名
            imagePDFName = imagePDFName.replaceAll(" ", "");
            String imgFolderPath ;
            if (dstImgFolder == null || dstImgFolder.equals("")) {
                imgFolderPath = imgPDFPath + File.separator + imagePDFName;// 获取图片存放的文件夹路径
            } else {
                imgFolderPath = dstImgFolder + File.separator + imagePDFName;
            }

            if (createDirectory(imgFolderPath)) {
                pdDocument = PDDocument.load(file);
                PDFRenderer renderer = new PDFRenderer(pdDocument);
				/* dpi越大转换后越清晰，相对转换速度越慢 */
                PdfReader reader = new PdfReader(PdfFilePath);
                int pages = reader.getNumberOfPages();
                StringBuffer imgFilePath = null;

                for (int i = 0; i < pages; i++) {
                    logger.info("转换第"+(i+1)+"张pdf文件");
                    String imgFilePathPrefix = imgFolderPath + File.separator + imagePDFName;
                    imgFilePath = new StringBuffer();
                    imgFilePath.append(imgFilePathPrefix);
                    imgFilePath.append("_");
                    imgFilePath.append(String.valueOf(i + 1));
                    imgFilePath.append(".png");
                    File dstFile = new File(imgFilePath.toString());
                    BufferedImage image = renderer.renderImageWithDPI(i, dpi);
                    ImageIO.write(image, "png", dstFile);
                    imageList.add(imgFilePath.toString());
                }
                pdDocument.close();
                logger.info("PDF文档转PNG图片成功！");

            } else {
                logger.info("PDF文档转PNG图片失败：" + "创建" + imgFolderPath + "失败");
            }

        } catch (IOException e) {
            e.printStackTrace();

        }
        return imageList;
    }

    public static List<String>  convertImage(String PdfFilePath, String dstImgFolder) {
        return convertImage(PdfFilePath, dstImgFolder, DEFAULT_DPI);
    }



    private static boolean createDirectory(String folder) {
        File dir = new File(folder);
        if (dir.exists()) {
            return true;
        } else {
            return dir.mkdirs();
        }
    }

    public static void main(String[] args) {
        convertImage("C:\\Users\\fangh\\Desktop\\test\\071802-java-2pic\\pdf\\【模板】账益达服务协议 王慧方.pdf", "C:\\Users\\fangh\\Desktop\\test\\071802-java-2pic\\png", DEFAULT_DPI);
    }


}

