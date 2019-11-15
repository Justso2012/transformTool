package per.alex.transform;


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import per.alex.transform.helper.DocumentHelper;
import per.alex.transform.helper.PdfHelper;

import java.io.File;

public class Deal {


    private static final String PNG_BASE_PATH = "C:\\Users\\fangh\\Desktop\\base_file\\png";
    //    public static final String PNG_BASE_PATH = "C:\\Users\\fangh\\Desktop\\test\\071802-java-2pic\\png";
//    public static final String PDF_BASE_PATH = "C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic\\签约客户word合同-pdf";
    private static final String PDF_BASE_PATH = "C:\\Users\\fangh\\Desktop\\base_file\\pdf";
    //    public static final String PDF_BASE_PATH = "C:\\Users\\fangh\\Desktop\\test\\071802-java-2pic\\pdf";
    private static final Logger logger = LoggerFactory.getLogger(Deal.class);

    private static final String DOC_TO_PDF_FOLDER = "C:\\Users\\fangh\\Desktop\\base_file\\pdf";
    private static final String DOC_FOLDER = "C:\\Users\\fangh\\Desktop\\base_file\\word";

    private void cyclePdfToPng() {
        File dir = new File(PDF_BASE_PATH);
        File[] files = dir.listFiles();
        int flag = 0;
        for (File file : files) {

            String fileName = file.getName();
            int index = fileName.lastIndexOf(".");
            String suffix = fileName.substring(index + 1);
            if (suffix.toLowerCase().equals("pdf")) {
                logger.info("----------开始转换合同:{}--------", file.getName());
                PdfHelper.convertImage(file.getPath(), PNG_BASE_PATH);

            }

        }
    }

    private void cycleDocToPdf() {
        File dir = new File(DOC_FOLDER);
        File[] files = dir.listFiles();
        int index = 0;
        for (File file : files) {

            logger.info("------开始word转pdf:{}------", file.getName());
            DocumentHelper.convertToPdf(file.getPath(), DOC_TO_PDF_FOLDER);
        }
    }


    public static void main(String[] args) {
        new Deal().cycleDocToPdf();
        new Deal().cyclePdfToPng();
//        new Deal().cycleDocToPdf();
    }


}
