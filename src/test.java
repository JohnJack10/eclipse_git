
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Scanner;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
public class test {
    public static void main(String[] args) {
//    	try {
//    		 
//            //读取word文档
//            XWPFDocument document = null;
//            try (InputStream in = Files.newInputStream(Paths.get("D:\\test.docx"))) {
//                document = new XWPFDocument(in);
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
// 
//            //将word转成pdf
//            PdfOptions options = PdfOptions.create();
//            try (OutputStream outPDF = Files.newOutputStream(Paths.get("D:\\test.pdf"))) {
//                PdfConverter.getInstance().convert(document, outPDF, options);
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//
//    }
    	jacobExcelToPDF("C:\\Users\\Administrator\\Desktop\\1.xlsx", "C:\\Users\\Administrator\\Desktop\\q.pdf");
    }
    /**
     * 使用jacob实现excel转PDF
     *
     * @param inputFilePath 导入Excel文件路径
     * @param outputFilePath 导出PDF文件路径
     */
    public static void jacobExcelToPDF(String inputFilePath, String outputFilePath) {
        ActiveXComponent ax = null;
        Dispatch excel = null;

        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));
		 	//禁用宏
            ax.setProperty("AutomationSecurity", new Variant(3));

            Dispatch excels = ax.getProperty("Workbooks").toDispatch();

            Object[] obj = {
                    inputFilePath,
                    new Variant(false),
                    new Variant(false)
            };

            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

			//转换格式
            Object[] obj2 = {
                    //PDF格式等于0
                    new Variant(0),
                    outputFilePath,
                    //0=标准（生成的PDF图片不会模糊），1=最小的文件
                    new Variant(0)
            };

            Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, obj2, new int[1]);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (excel != null) {
                Dispatch.call(excel, "Close", new Variant(false));
            }
            if (ax != null) {
                ax.invoke("Quit", new Variant[]{});
                ax = null;
            }
            ComThread.Release();
        }

    }
}