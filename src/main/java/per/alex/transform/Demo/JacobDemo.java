package per.alex.transform.Demo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;

public class JacobDemo {

    private static final Logger logger = LoggerFactory.getLogger(JacobDemo.class);

    private static final int wdFormatPDF = 17; //pdf格式

    private void wordToPdf(String sourcePath, String destPath) {
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            doc = Dispatch.call(docs, "Open", sourcePath).toDispatch();
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

    private static void main(String[] args) {
//        new JacobDemo().wordToPdf("C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic\\demo.docx", "C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic\\demo.pdf");
        new JacobDemo().wordToPdf("C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic\\（厦门天网时代信息科技有限公司）厦门链上云科技有限公司.docx",
                "C:\\Users\\fangh\\Desktop\\test\\073101-java-word2pic\\（厦门天网时代信息科技有限公司）厦门链上云科技有限公司.pdf");
    }
}
