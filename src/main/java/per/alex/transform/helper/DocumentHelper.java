package per.alex.transform.helper;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;

public class DocumentHelper {

    private static final int wdFormatPDF = 17; //pdf格式

    /**
     * word转换为pdf
     *
     * @param documentPath word绝对路径
     * @param destPdfFoler pdf存放文件夹
     */
    public static void convertToPdf(String documentPath, String destPdfFoler) {
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            File file = new File(documentPath);
            String fileName = file.getName();
            int index = fileName.lastIndexOf(".");
            String suffixName = fileName.substring(index + 1);
            String name = fileName.substring(0, index);

            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            doc = Dispatch.call(docs, "Open", documentPath).toDispatch();
            String destPath = destPdfFoler + File.separator + name + ".pdf";
            File destFile = new File(destPath);
            if (destFile.exists()) {
                destFile.delete();
            }
            Dispatch.call(doc, "SaveAs", destPath, wdFormatPDF);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Dispatch.call(doc, "Close", false);
            if (app != null) {
                app.invoke("Quit", new Variant[]{});
            }
        }
        ComThread.Release();
    }

    public static void main(String[] args) {
        convertToPdf("\\\\172.16.19.83\\download\\签约客户word合同\\签约客户word合同\\【道恒财务管理】账益达服务协议（原渠道客户）_20180903.docx",
                "C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic");
    }
}
