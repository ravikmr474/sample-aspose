package com.aspose.demo.asposeimagecapture;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;

public class FileGenerator {

	public ByteArrayOutputStream createDocxFile() throws FileNotFoundException, Exception {

		License license = new License();
		InputStream licenseFile = FileGenerator.class.getClassLoader().getResourceAsStream("Aspose.Total.Java.lic");
		license.setLicense(licenseFile);
		
		InputStream docxFile = FileGenerator.class.getClassLoader().getResourceAsStream("sample.docx");
		Document doc = new Document(docxFile);

		InputStream excelFile = FileGenerator.class.getClassLoader().getResourceAsStream("my-excel-demo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
//		XSSFSheet spreadsheet = workbook.getSheet("sheet1");
//		spreadsheet.setMargin(Sheet.LeftMargin, 0.0);
//		spreadsheet.setMargin(Sheet.RightMargin, 0.0);
//		spreadsheet.setMargin(Sheet.TopMargin, 0.0);
//		spreadsheet.setMargin(Sheet.BottomMargin, 0.0);
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		try {
			DocumentBuilder builder = new DocumentBuilder(doc);
//			builder.writeln("sheet_1_image");
//			File fileImage = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Output12345.png");
//	        FileInputStream imageData
//	            = new FileInputStream(fileImage);
	        Map<String, BufferedImage> imageURLMap = FileMerge.captureImageFromExcelSpire(workbook);
//			builder.insertImage(imageURLMap.get("sheet_1_image"),468.0,450.0);
	        builder.insertImage(imageURLMap.get("sheet_1_image"));
			builder.insertImage(imageURLMap.get("sheet_2_image"));
//			builder.insertImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Output12345.png");
			
//			builder.writeln("sheet_2_image");
//			InputStream inputStream = FileMerge.writeImage(convertToInputStream(doc), workbook);
			InputStream inputStream =convertToInputStream(doc);
			byte[] buffer = new byte[1000];
			try {
				int temp;

				while ((temp = inputStream.read(buffer)) != -1) {
					byteArrayOutputStream.write(buffer, 0, temp);
				}
			} catch (Exception e) {
				System.out.println(e);
			}

		} catch (Exception e1) {
			System.out.println("Exception: " + e1.getMessage());
			e1.printStackTrace();
		}
		return byteArrayOutputStream;

	}
	
	public ByteArrayOutputStream createDocxFileReport(String file) throws FileNotFoundException, Exception {

		License license = new License();
		InputStream licenseFile = FileGenerator.class.getClassLoader().getResourceAsStream("Aspose.Total.Java.lic");
		license.setLicense(licenseFile);
		
		InputStream docxFile = FileGenerator.class.getClassLoader().getResourceAsStream("sample.docx");
		Document doc = new Document(docxFile);

		InputStream excelFile = null;
		if("modified".equalsIgnoreCase(file)) {
			excelFile = FileGenerator.class.getClassLoader().getResourceAsStream("my-excel-modified.xlsx");
		} else {
			excelFile = FileGenerator.class.getClassLoader().getResourceAsStream("my-excel-demo.xlsx");
		}
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		try {
			DocumentBuilder builder = new DocumentBuilder(doc);
	        Map<String, InputStream> imageURLMap = FileMerge.captureImageFromExcelAspose(workbook);
	        builder.insertImage(imageURLMap.get("image_sheet_1"));
			
			InputStream inputStream =convertToInputStream(doc);
			byte[] buffer = new byte[1000];
			try {
				int temp;

				while ((temp = inputStream.read(buffer)) != -1) {
					byteArrayOutputStream.write(buffer, 0, temp);
				}
			} catch (Exception e) {
				System.out.println(e);
			}

		} catch (Exception e1) {
			System.out.println("Exception: " + e1.getMessage());
			e1.printStackTrace();
		}
		return byteArrayOutputStream;

	}

	public static InputStream convertToInputStream(Document doc) throws Exception {
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		doc.save(outStream, SaveFormat.DOCX);
		byte[] docBytes = outStream.toByteArray();
		InputStream inputStream = new ByteArrayInputStream(docBytes);
		return inputStream;
	}

}
