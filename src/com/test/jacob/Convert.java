package com.test.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

/**
 * Created by jichao on 2017/12/29.
 */
public class Convert {

    public static boolean word2PDF(String inputFile, String pdfFile) {
        try {
            // 打开word应用程序
            ActiveXComponent app = new ActiveXComponent("Word.Application");
            // 设置word不可见
            app.setProperty("Visible", false);
            // 获得word中所有打开的文档,返回Documents对象
            Dispatch docs = app.getProperty("Documents").toDispatch();
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            Dispatch doc = Dispatch.call(docs, "Open", inputFile, false, true)
                    .toDispatch();
            // 调用Document对象的SaveAs方法，将文档保存为pdf格式
            /*
             * Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF
             * //word保存为pdf格式宏，值为17 );
             */
            Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, 17);// word保存为pdf格式宏，值为17
            // 关闭文档
            Dispatch.call(doc, "Close", false);
            // 关闭word应用程序
            app.invoke("Quit", 0);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
    // excel转换为pdf
    public static boolean excel2PDF(String inputFile, String pdfFile) {
        try {
            ActiveXComponent app = new ActiveXComponent("Excel.Application");
            app.setProperty("Visible", false);
            Dispatch excels = app.getProperty("Workbooks").toDispatch();
            Dispatch excel = Dispatch.call(excels, "Open", inputFile, false,
                    true).toDispatch();
            Dispatch.call(excel, "ExportAsFixedFormat", 0, pdfFile);
            Dispatch.call(excel, "Close", false);
            app.invoke("Quit");
            return true;
        } catch (Exception e) {
            return false;
        }
    }
    public static boolean ppt2PDF(String inputFile, String pdfFile) {
        try {
            ActiveXComponent app = new ActiveXComponent(
                    "PowerPoint.Application");
            // app.setProperty("Visible", msofalse);
            Dispatch ppts = app.getProperty("Presentations").toDispatch();

            Dispatch ppt = Dispatch.call(ppts, "Open", inputFile, true,// ReadOnly
                    true,// Untitled指定文件是否有标题
                    false// WithWindow指定文件是否可见
            ).toDispatch();

            Dispatch.call(ppt, "SaveAs", pdfFile, 32);

            Dispatch.call(ppt, "Close");

            app.invoke("Quit");
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public static void main(String[] args) {
        Convert.word2PDF("C:\\Users\\jichao\\Desktop\\4.docx", "C:\\Users\\jichao\\Desktop\\444.pdf");

//        Convert.excel2PDF("C:\\Users\\jichao\\Desktop\\4.ppt", "C:\\Users\\jichao\\Desktop\\444.pdf");
//
//        Convert.ppt2PDF("C:\\Users\\jichao\\Desktop\\4.ppt", "C:\\Users\\jichao\\Desktop\\444.pdf");

        System.out.println("转换完成！");
    }

}
