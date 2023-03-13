package com.aspose.demo.asposeimagecapture;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Service;

@Service
public class ExportService {

	public void download(HttpServletResponse response) throws Exception {
		String fileName = "Report.zip";
		String zipDirName = "/tmp/" + fileName;
		FileOutputStream fileOutputStream = new FileOutputStream(zipDirName);
		ZipOutputStream zipOutputStream = new ZipOutputStream(fileOutputStream);
		zipOutputStream.flush();
		try {

			ZipEntry docx = new ZipEntry("document.docx");
			zipOutputStream.putNextEntry(docx);
			FileGenerator fsDocxGene = new FileGenerator();
			ByteArrayOutputStream docxFileOutputStream = fsDocxGene.createDocxFile();

			docxFileOutputStream.writeTo(zipOutputStream);
			zipOutputStream.closeEntry();
			zipOutputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		doDownload(zipDirName, response);
	}

	public static void doDownload(String filePath, HttpServletResponse response) throws IOException {

		String fullPath = filePath;
		File downloadFile = new File(fullPath);
		FileInputStream inputStream = new FileInputStream(downloadFile);

		response.setContentType("application/octet-stream");
		response.setContentLength((int) downloadFile.length());

		String headerKey = "Content-Disposition";
		String headerValue = String.format("attachment; filename=\"%s\"", downloadFile.getName());
		response.setHeader(headerKey, headerValue);

		OutputStream outStream = response.getOutputStream();

		byte[] buffer = new byte[4096];
		int bytesRead = -1;

		while ((bytesRead = inputStream.read(buffer)) != -1) {
			outStream.write(buffer, 0, bytesRead);
		}
		inputStream.close();
		outStream.close();

	}
}
