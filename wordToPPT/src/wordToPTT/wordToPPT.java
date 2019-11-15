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

		String content = null; // ���e
		int pages = 0;// ����
		try {

			// Ū��word�ɮ�
			OPCPackage file = POIXMLDocument.openPackage(filePath);
			XWPFDocument docx = new XWPFDocument(file);
			@SuppressWarnings("resource")
			POIXMLTextExtractor readFile = new XWPFWordExtractor(docx);

			// ���o���e
			content = readFile.getText();

			// ���o����
			pages = docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();

//			System.out.println(content);
			System.out.println("����" + pages + "��");
			// �_�Ʀ۰ʷs�W�ťխ��ɦ�����
			if (pages % 2 != 0) {
				// �q�������Ы�XWPFRun
				XWPFParagraph p1 = docx.createParagraph();
				XWPFRun run = p1.createRun();

				// �]�w�r��20��20��(�s�W�@��)
				run.setFontSize(20);
				for (int i = 0; i < 20; i++) {
					run.addCarriageReturn();

				}

				// ��X��WORD��
				fileName = fileName + "Edit";
				System.out.println("���Ƴ渹,�ഫ������,�s�W�ťխ�,�নword�ɮצW��:" + fileName);
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

	// �নPDF
//	public static void transToPDF(String filePath, String fileName) {
//		
//		FileOutputStream os = null;
//		try {
//			File PDFfile = new File("D:/" + fileName + ".pdf"); // �s�ؤ@�Ӫť�pdf���
//			os = new FileOutputStream(PDFfile);
//			
//			Document document = new Document("D:/" + fileName + ".docx"); // Address�O�N�n�Q��ƪ�word���
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
	//�নPDF
	public static int transToPDF(String filePath, String outputFile) throws Exception {

		System.out.println("�Ұ�Word");
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("Word.Application");
			// �]�wword���i��
			app.setProperty("Visible", new Variant(false));
			// �}��word�ɮ�
			Dispatch docs = app.getProperty("Documents").toDispatch();
//	          doc = Dispatch.call(docs,  "Open" , sourceFile).toDispatch();   
			doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
					new Object[] { filePath, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
			System.out.println("�}�Ҥ��" + filePath);
			System.out.println("�ഫ����" + outputFile);
			File tofile = new File(outputFile);
			// System.err.println(getDocPageSize(new File(sfileName)));
			if (tofile.exists()) {
				tofile.delete();
			}
//	          Dispatch.call(doc, "SaveAs",  destFile,  17);   
			// �@��html�榡�x�s���{���ɮסG�G�޼� new
			// Variant(8)�䤤8���word��html;7���word��txt;44���Excel��html;17���word�নpdf�C�C
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { outputFile, new Variant(17) }, new int[1]);
			long end = System.currentTimeMillis();
			System.out.println("�ഫ���� �ήɡG" + (end - start) + "ms.");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("========Error:����ഫ���ѡG" + e.getMessage());
		} catch (Throwable t) {
			t.printStackTrace();
		} finally {
			// ����word
//			Dispatch.call(doc, "Close", false);
			System.out.println("�������");
			if (app != null)
				app.invoke("Quit", new Variant[] {});
		}
		// close winword.exe�{��
		ComThread.Release();
		return 1;
	}

}
