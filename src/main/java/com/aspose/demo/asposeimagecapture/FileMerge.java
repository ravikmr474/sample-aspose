package com.aspose.demo.asposeimagecapture;

import java.awt.Dimension;
import java.awt.HeadlessException;
import java.awt.RenderingHints;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import com.amazonaws.auth.AWSCredentials;
import com.amazonaws.auth.AWSStaticCredentialsProvider;
import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.regions.Regions;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3ClientBuilder;
import com.amazonaws.services.s3.model.ObjectMetadata;
import com.amazonaws.services.s3.model.PutObjectRequest;
import com.amazonaws.services.s3.model.S3Object;
import com.amazonaws.services.s3.model.S3ObjectInputStream;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.FontConfigs;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.Range;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Style;
import com.aspose.cells.TiffCompression;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.words.FolderFontSource;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import com.aspose.words.SystemFontSource;
import com.spire.xls.PageSetup;
import com.spire.xls.core.spreadsheet.HTMLOptions;

public class FileMerge {

	public static InputStream writeImage(InputStream inputStream, XSSFWorkbook workbook) throws IOException {

		Map<String, OutputStream> imageURLMap = takeSnapshotFromExcelSpire(workbook);
		XWPFDocument doc = new XWPFDocument(inputStream);
		try {
			for (XWPFParagraph p : doc.getParagraphs()) {
				String text2 = p.getText().trim();
					List<XWPFRun> runs = p.getRuns();
					if (runs != null) {
						for (XWPFRun r : runs) {
							String text = r.getText(0);
							if (text != null) {
								BufferedImage image;
								ByteArrayOutputStream bout;
								ByteArrayInputStream bin;
								InputStream imageInputStream = getInputStream(imageURLMap.get(text2));
								image = ImageIO.read(imageInputStream);
								Dimension dim = new Dimension(image.getWidth(), image.getHeight());
								// Dimension width
								double width = dim.getWidth();
								double height = dim.getHeight();
								double scaling = 1.0;
								if (width > 72 * 6.5)
									scaling = (72 * 6.5) / width;
								bout = new ByteArrayOutputStream();
								ImageIO.write(image, "png", bout);
								bout.flush();
								bin = new ByteArrayInputStream(bout.toByteArray());
								r.addPicture(bin, XWPFDocument.PICTURE_TYPE_PNG, "", Units.toEMU(width*scaling),
										Units.toEMU(height*scaling));
								r.getCTR().getDrawingArray(0).getInlineArray(0).addNewCNvGraphicFramePr()
										.addNewGraphicFrameLocks().setNoChangeAspect(true);
								text = text.replace(text2, "");
								r.setText(text, 0);
							}
							break;
						}
					}
			}
		} catch (Exception e) {
			System.out.println("Exception: " + e);
			e.printStackTrace();
		}
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		doc.write(outStream);
		byte[] docBytes = outStream.toByteArray();

		return new ByteArrayInputStream(docBytes);

	}

	public static InputStream getInputStream(OutputStream outputStream) {
		InputStream isFromFirstData = new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
		return isFromFirstData;
	}

	public static Map<String, OutputStream> takeSnapshotFromExcel(XSSFWorkbook workbookFile) {

		Map<String, OutputStream> imageOutputStreamMap = new HashMap<String, OutputStream>();
		try {

//			ByteArrayOutputStream bos = new ByteArrayOutputStream();
//			workbookFile.write(bos);
//			byte[] barray = bos.toByteArray();
//			InputStream is = new ByteArrayInputStream(barray);
//
//			Workbook workbook = new Workbook(is);
//			FontConfigs.setDefaultFontName("Arial");
//			workbook.calculateFormula();
//			Style defaultStyle = workbook.getDefaultStyle();
//			FontSettings fs = FontSettings.getDefaultInstance();
//			fs.setFontsSources(
//					new FontSourceBase[] { new SystemFontSource(), new FolderFontSource("/usr/share/fonts", true) });
//
//			defaultStyle.getFont().setName("Arial");
//			defaultStyle.getFont().setSize(10);
//			defaultStyle.setTextWrapped(true);
//			defaultStyle.getFont().setBold(false);
//			workbook.setDefaultStyle(defaultStyle);
//			WorksheetCollection worksheets = workbook.getWorksheets();
//
//			int sheetCount = worksheets.getCount();
			
			Workbook workbook = new Workbook("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\My-excel.xlsx");
			FontConfigs.setDefaultFontName("Arial");
			workbook.calculateFormula();
			Style defaultStyle = workbook.getDefaultStyle();

			defaultStyle.getFont().setName("Arial");
			defaultStyle.getFont().setSize(10);
			defaultStyle.setTextWrapped(true);
			defaultStyle.getFont().setBold(false);
			workbook.setDefaultStyle(defaultStyle);
			
			workbook.getWorksheets().get(0).getPageSetup().setPrintArea("A1:A27:B1:B27");
			workbook.getWorksheets().get(0).getPageSetup().setLeftMargin(0);
			workbook.getWorksheets().get(0).getPageSetup().setRightMargin(0);
			workbook.getWorksheets().get(0).getPageSetup().setTopMargin(0);
			workbook.getWorksheets().get(0).getPageSetup().setBottomMargin(0);
			
			ImageOrPrintOptions options = new ImageOrPrintOptions();
			options.setOnePagePerSheet(true);
			options.setImageType(ImageType.PNG);
//			options.setImageType(ImageType.JPEG);
			options.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			options.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
					RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
			options.setHorizontalResolution(100);
			options.setVerticalResolution(100);
			options.setTiffCompression(TiffCompression.COMPRESSION_LZW);
			options.setQuality(100);
			options.setCheckWorkbookDefaultFont(true);
			options.setDefaultFont("Arial");
			
			SheetRender sr = new SheetRender(workbook.getWorksheets().get(0), options);
			ByteArrayOutputStream outStream = new ByteArrayOutputStream();
			sr.toImage(0, "C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\output.png");
			sr.toImage(0, outStream);
			imageOutputStreamMap.put("sheet_1_image", outStream);
//			for (int index = 0; index < sheetCount; index++) {
//				Worksheet workbookName = workbook.getWorksheets().get(index);
//				workbookName.getPageSetup().setLeftMargin(0);
//				workbookName.getPageSetup().setRightMargin(0);
//				workbookName.getPageSetup().setTopMargin(0);
//				workbookName.getPageSetup().setBottomMargin(0);
//				String name = workbookName.getName();
//				Range range = workbookName.getCells().getMaxDisplayRange();
//				if (range != null) {
//					int tcols = range.getColumnCount();					
//					ImageOrPrintOptions options = new ImageOrPrintOptions();
//					options.setOnePagePerSheet(true);
//					options.setImageType(ImageType.EMF);
////					options.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
////					options.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
////							RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
//					options.setHorizontalResolution(100);
//					options.setVerticalResolution(100);
////					options.setTiffCompression(TiffCompression.COMPRESSION_LZW);
////					options.setQuality(100);
//					options.setCheckWorkbookDefaultFont(true);
////					options.setDefaultFont("Arial");
//					options.setOnlyArea(true);
//					SheetRender sr = null;
//					Cells cells = workbookName.getCells();
//					Cell lastCell = cells.getLastCell();
//					int row = lastCell.getRow();
//					String imagePrintArea = "";
//					ByteArrayOutputStream outStream = null;
//					switch (name) {
//					case "sheet_1_image":
//						try {
//							imagePrintArea = "A5:A31:B5:B31";
//							workbookName.getPageSetup().setPrintArea(imagePrintArea);
//							sr = new SheetRender(workbookName, options);
//							outStream = new ByteArrayOutputStream();
//							sr.toImage(0, "C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\aspose-lib\\sheet_image.emf");
//							imageOutputStreamMap.put("sheet_1_image", outStream);
//						} catch (Exception e) {
//							e.printStackTrace();
//						}
//						break;
//					}
//				}
//			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		return imageOutputStreamMap;
	}

	public static Map<String, OutputStream> takeSnapshotFromExcelSpire(XSSFWorkbook workbookFile) throws IOException {
		
//		com.spire.xls.Workbook workbook = new com.spire.xls.Workbook();	
//		ByteArrayOutputStream bos = new ByteArrayOutputStream();
//		workbookFile.write(bos);
//		byte[] barray = bos.toByteArray();
//		InputStream is = new ByteArrayInputStream(barray);
//		workbook.loadFromStream(is);
//
//		com.spire.xls.Worksheet sheet = workbook.getWorksheets().get(0);
//
//		HTMLOptions options = new HTMLOptions();	
//		ByteArrayOutputStream outputStreamHTML = new ByteArrayOutputStream();
//		sheet.saveToHtml(outputStreamHTML, options);
////		sheet.saveToHtml("src/main/resources/htmlFile.html", options);
//		try {
//			uploadFile("htmlFile.html", outputStreamHTML);
//			downloadFile();
//		} catch (IOException e1) {
//			e1.printStackTrace();
//		}
		
//		System.setProperty("webdriver.chrome.driver", "src/main/resources/chromedriver.exe"); // windows
		System.setProperty("webdriver.chrome.driver", "/usr/bin/chromedriver"); // linux
		
		ChromeOptions options1 = new ChromeOptions();
//		options1.setHeadless(true);
		options1.addArguments("--remote-allow-origins=*");
		options1.addArguments("--headless");
//		options1.addArguments("start-maximized"); // open Browser in maximized mode
//		options1.addArguments("disable-infobars"); // disabling infobars
//		options1.addArguments("--disable-extensions"); // disabling extensions
//		options1.addArguments("--disable-gpu"); // applicable to windows os only
		options1.addArguments("--disable-dev-shm-usage"); // overcome limited resource problems
		options1.addArguments("--no-sandbox"); // Bypass OS security model
		WebDriver driver = null;
		try {
			driver = new ChromeDriver(options1);
			driver.get("https://s3.amazonaws.com/seamlessserver/signinlogo/fsimages/htmlFile.html");
			driver.manage().window().maximize();
		} catch (Exception e) {
			System.out.println("MEsage:-- "+e.getMessage());
			e.printStackTrace();
		}
		
//		int height = 814;
//		try {
//			width = (int) Toolkit.getDefaultToolkit().getScreenSize().getWidth();
//			height = (int) Toolkit.getDefaultToolkit().getScreenSize().getHeight()-50;
//		} catch (HeadlessException e) {
//			e.printStackTrace();
//		}
        
        // Get the entire page screenshot
        File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        FileUtils.copyFile(screenshot, outputStream);
        InputStream isFromFirstData = new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
        
        BufferedImage image = ImageIO.read(isFromFirstData);
        
        
        BufferedImage croppedImage = image.getSubimage(10, 70, 696, 500);	
        ByteArrayOutputStream baos = new ByteArrayOutputStream();;
        ImageIO.write(croppedImage, "png", baos);
        Map<String, OutputStream> imageOutputStreamMap = new HashMap<String, OutputStream>();
        imageOutputStreamMap.put("sheet_1_image", baos);
        System.out.println("Done...done again");
        driver.quit();
		return imageOutputStreamMap;
	}
	
}