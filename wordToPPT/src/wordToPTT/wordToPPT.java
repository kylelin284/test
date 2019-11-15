package wordToPTT;

import java.io.File;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;

import com.aspose.pdf.Document;
import com.aspose.pdf.SaveFormat;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
//
public class wordToPPT {
	public static void main(String[] args) {
		String fileName = "testFile";
		String filePath = "D:/" + fileName + ".docx";
		readWord(filePath, fileName);
	}

	public static void readWord(String filePath, String fileName) {

		String content = null; // 內容
		int pages = 0;// 頁數
		try {

			// 讀取word檔案
			OPCPackage file = POIXMLDocument.openPackage(filePath);
			XWPFDocument docx = new XWPFDocument(file);
			@SuppressWarnings("resource")
			POIXMLTextExtractor readFile = new XWPFWordExtractor(docx);

			// 取得內容
			content = readFile.getText();

			// 取得頁數
			pages = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();

//			System.out.println(content);
			System.out.println("頁數" + pages + "頁");
			// 奇數自動新增空白頁補成偶數
			if (pages % 2 != 0) {
				// 段落末尾創建XWPFRun
				XWPFParagraph p1 = docx.createParagraph();
				XWPFRun run = p1.createRun();

				// 設定字體20換20行(新增一頁)
				run.setFontSize(20);
				for (int i = 0; i < 20; i++) {
					run.addCarriageReturn();

				}

				// 輸出成WORD檔
				fileName = fileName + "Edit";
				System.out.println("頁數單號,轉換成雙號,新增空白頁,轉成word檔案名稱:" + fileName);
				FileOutputStream out = new FileOutputStream("D:/" + fileName + ".docx");
				docx.write(out);
				out.close();
				try {
					transToPDF("D:/" + fileName + ".docx", "D:/" + fileName + ".pdf");
				} catch (Exception e) {
					e.printStackTrace();
				}
			} else {
				try {
					transToPDF(filePath, "D:/" + fileName + ".pdf");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

			file.close();
		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	// 轉成PDF
//	public static void transToPDF(String filePath, String fileName) {
//		
//		FileOutputStream os = null;
//		try {
//			File PDFfile = new File("D:/" + fileName + ".pdf"); // 新建一個空白pdf文件
//			os = new FileOutputStream(PDFfile);
//			
//			Document document = new Document("D:/" + fileName + ".docx"); // Address是將要被轉化的word文件
//			
//			document.save(os, SaveFormat.Pdf);
//		} catch (Exception e) {
//			e.printStackTrace();
//		} finally {
//			if (os != null) {
//				try {
//					os.close();
//				} catch (IOException e) {
//					e.printStackTrace();
//				}
//			}
//		}
//	}
	//轉成PDF
	public static int transToPDF(String filePath, String outputFile) throws Exception {

		System.out.println("啟動Word");
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("Word.Application");
			// 設定word不可見
			app.setProperty("Visible", new Variant(false));
			// 開啟word檔案
			Dispatch docs = app.getProperty("Documents").toDispatch();
//	          doc = Dispatch.call(docs,  "Open" , sourceFile).toDispatch();   
			doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
					new Object[] { filePath, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			System.out.println("開啟文件" + filePath);
			System.out.println("轉換文件到" + outputFile);
			File tofile = new File(outputFile);
			// System.err.println(getDocPageSize(new File(sfileName)));
			if (tofile.exists()) {
				tofile.delete();
			}
//	          Dispatch.call(doc, "SaveAs",  destFile,  17);   
			// 作為html格式儲存到臨時檔案：：引數 new
			// Variant(8)其中8表示word轉html;7表示word轉txt;44表示Excel轉html;17表示word轉成pdf。。
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { outputFile, new Variant(17) }, new int[1]);
			long end = System.currentTimeMillis();
			System.out.println("轉換完成 用時：" + (end - start) + "ms.");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("========Error:文件轉換失敗：" + e.getMessage());
		} catch (Throwable t) {
			t.printStackTrace();
		} finally {
			// 關閉word
//			Dispatch.call(doc, "Close", false);
			System.out.println("關閉文件");
			if (app != null)
				app.invoke("Quit", new Variant[] {});
		}
		// close winword.exe程序
		ComThread.Release();
		return 1;
	}

}
