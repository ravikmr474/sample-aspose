package com.aspose.demo.asposeimagecapture;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

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

		InputStream excelFile = FileGenerator.class.getClassLoader().getResourceAsStream("My-excel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
//		XSSFSheet spreadsheet = workbook.getSheet("sheet1");
//		spreadsheet.setMargin(Sheet.LeftMargin, 0.0);
//		spreadsheet.setMargin(Sheet.RightMargin, 0.0);
//		spreadsheet.setMargin(Sheet.TopMargin, 0.0);
//		spreadsheet.setMargin(Sheet.BottomMargin, 0.0);
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		try {
			DocumentBuilder builder = new DocumentBuilder(doc);
			builder.writeln("sheet_1_image");
			
			builder.writeln("sheet_2_image");
			InputStream inputStream = FileMerge.writeImage(convertToInputStream(doc), workbook);
			int available = inputStream.available();
			System.out.println("available:-  "+available);
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
