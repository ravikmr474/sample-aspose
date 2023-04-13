package com.aspose.demo.asposeimagecapture;

import java.awt.Dimension;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Rectangle;
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
import com.aspose.cells.FontConfigs;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Style;
import com.aspose.cells.TiffCompression;
import com.aspose.cells.Workbook;
import com.spire.xls.PageOrientationType;
import com.spire.xls.PageSetup;
import com.spire.xls.PaperSizeType;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.HTMLOptions;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.coordinates.Coords;
import ru.yandex.qatools.ashot.coordinates.WebDriverCoordsProvider;
import ru.yandex.qatools.ashot.cropper.ImageCropper;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

public class FileMerge {

//	public static InputStream writeImage(InputStream inputStream, XSSFWorkbook workbook) throws IOException {

//		Map<String, BufferedImage> imageURLMap = captureImageFromExcelSpire(workbook);
//		XWPFDocument doc = new XWPFDocument(inputStream);
//		try {
//			for (XWPFParagraph p : doc.getParagraphs()) {
//				String text2 = p.getText().trim();
//					List<XWPFRun> runs = p.getRuns();
//					if (runs != null) {
//						for (XWPFRun r : runs) {
//							String text = r.getText(0);
//							if (text != null) {
//								BufferedImage image;
//								ByteArrayOutputStream bout;
//								ByteArrayInputStream bin;
//								InputStream imageInputStream = getInputStream(imageURLMap.get(text2));
//								
////								image = ImageIO.read(imageInputStream);
////								Dimension dim = new Dimension(image.getWidth(), image.getHeight());
////								// Dimension width
////								double width = dim.getWidth();
////								double height = dim.getHeight();
////								double scaling = 1.0;
////								if (width > 72 * 6.5)
////									scaling = (72 * 6.5) / width;
////								bout = new ByteArrayOutputStream();
////								ImageIO.write(image, "png", bout);
////								bout.flush();
////								bin = new ByteArrayInputStream(bout.toByteArray());
//								double	width = 468.0;
//								double	height = 450.0;
//								System.out.println("width:- "+width+" height:- "+height);
//								 File fileImage = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Output12345.png");
//							        FileInputStream imageData
//							            = new FileInputStream(fileImage);
//								r.addPicture(imageData, XWPFDocument.PICTURE_TYPE_PNG, "", Units.toEMU(width),
//										Units.toEMU(height));
//								r.getCTR().getDrawingArray(0).getInlineArray(0).addNewCNvGraphicFramePr()
//										.addNewGraphicFrameLocks().setNoChangeAspect(true);
//								text = text.replace(text2, "");
//								r.setText(text, 0);
//							}
//							break;
//						}
//					}
//			}
//		} catch (Exception e) {
//			System.out.println("Exception: " + e);
//			e.printStackTrace();
//		}
//		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
//		doc.write(outStream);
//		byte[] docBytes = outStream.toByteArray();
//
//		return new ByteArrayInputStream(docBytes);

//	}

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
		
		com.spire.xls.Workbook workbook = new com.spire.xls.Workbook();	
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		workbookFile.write(bos);
		byte[] barray = bos.toByteArray();
		InputStream is = new ByteArrayInputStream(barray);
		workbook.loadFromStream(is);

		com.spire.xls.Worksheet sheet = workbook.getWorksheets().get(0);

		HTMLOptions options = new HTMLOptions();	
		ByteArrayOutputStream outputStreamHTML = new ByteArrayOutputStream();
//		sheet.saveToHtml(outputStreamHTML, options);
		sheet.saveToHtml("src/main/resources/htmlFile.html", options);
		
//		try {
//			uploadFile("htmlFile.html", outputStreamHTML);
//			downloadFile();
//		} catch (IOException e1) {
//			e1.printStackTrace();
//		}
		
//		System.setProperty("webdriver.chrome.driver", "src/main/resources/chromedriver.exe"); // windows
//		System.setProperty("webdriver.chrome.driver", "/usr/bin/chromedriver"); // linux
		
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
			System.out.println("Exception Message:-- "+e.getMessage());
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
        
        
        BufferedImage croppedImage = image.getSubimage(5, 50, 560, 550);
        
        File outputFile1 = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Output12345.png");
        ImageIO.write(croppedImage, "png", outputFile1);
        
        ByteArrayOutputStream baos = new ByteArrayOutputStream();;
        ImageIO.write(croppedImage, "png", baos);
        Map<String, OutputStream> imageOutputStreamMap = new HashMap<String, OutputStream>();
        imageOutputStreamMap.put("sheet_1_image", baos);
        System.out.println("Done...done again");
        driver.quit();
		return imageOutputStreamMap;
	}
	
	public static Map<String, BufferedImage> captureImageFromExcelSpire(XSSFWorkbook workbookFile) throws IOException {
		
		Map<String, BufferedImage> imageOutputStreamMap = new HashMap<String, BufferedImage>();
		
		com.spire.xls.Workbook workbook = new com.spire.xls.Workbook();	
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		workbookFile.write(bos);
		byte[] barray = bos.toByteArray();
		InputStream is = new ByteArrayInputStream(barray);
		workbook.loadFromStream(is);
		
//		String zipDirName = "/tmp/Financial_Report/";
		int sheetCount = workbook.getWorksheets().getCount();	
		HTMLOptions options = null;
		
//		System.setProperty("webdriver.chrome.driver", "src/main/resources/chromedriver.exe"); // windows
		System.setProperty("webdriver.chrome.driver", "/usr/bin/chromedriver"); // linux
		
		ChromeOptions options1 = new ChromeOptions();
		options1.addArguments("--remote-allow-origins=*");
		options1.addArguments("--headless");
		options1.addArguments("--disable-dev-shm-usage"); // overcome limited resource problems
		options1.addArguments("--no-sandbox"); // Bypass OS security model

		WebDriver driver = null;
		File screenshot = null;
		ByteArrayOutputStream outputStream = null;
		InputStream isFromFirstData = null;
		BufferedImage image = null;
		BufferedImage croppedImage = null;
		ByteArrayOutputStream baos = null;
		ByteArrayOutputStream outputStreamHTML = null;
//		workbook.getConverterSetting().setXDpi(320);
//	    workbook.getConverterSetting().setYDpi(320);
	    
		for(int count =0; count<sheetCount;count++) {			
			com.spire.xls.Worksheet worksheet = workbook.getWorksheets().get(count);
			File pngFile = null;
			switch (worksheet.getName()) {
			case "SOFC":
				
				options = new HTMLOptions();	
				outputStreamHTML = new ByteArrayOutputStream();
//				sheet.saveToHtml(outputStreamHTML, options);
				worksheet.saveToHtml("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOFCHtml.html", options);				
				
//				worksheet.saveToImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOFCExcelToPng.png");
				
//				worksheet.saveToImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOFCRangeImageWithRange.png",5, 1, 27, 2);
				
//				worksheet.saveToHtml(outputStreamHTML, options);
//				uploadFile("sheet1.html", outputStreamHTML);
				try {
					driver = new ChromeDriver(options1);
//					driver.get("https://seamlessserver.s3.amazonaws.com/signinlogo/fsimages/sheet1.html");
					driver.get("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\demo1.html");
					driver.manage().window().maximize();
				} catch (Exception e) { 
					System.out.println("Exception Message:-- "+e.getMessage());
					e.printStackTrace();
				}
						        
		        // Get the entire page screenshot
				WebElement findElement = driver.findElement(By.tagName("table"));
//				WebElement findElement = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td[1]/div"));
				
				
				Screenshot screenshotHeader = new AShot().coordsProvider(new WebDriverCoordsProvider())
						.shootingStrategy(ShootingStrategies.viewportPasting(1000)).takeScreenshot(driver, findElement);
			    try {
			        ImageIO.write(screenshotHeader.getImage(),"png", 
			        		new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\demoHtmlToImage1.png"));
			    } catch (IOException e) {
			        e.printStackTrace();
			    }
//		        screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
//				
//		        outputStream = new ByteArrayOutputStream();
//		        FileUtils.copyFile(screenshot, outputStream);
//		        isFromFirstData = new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
//		        
//		        image = ImageIO.read(isFromFirstData);
//		        croppedImage = image.getSubimage(rect.getX(), 50, rect.getWidth(), rect.getHeight()-40);		
		        
		        pngFile = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\demo1.png");
				ImageIO.write(screenshotHeader.getImage(), "png", pngFile);
					
//		        baos = new ByteArrayOutputStream();
//		        ImageIO.write(image, "png", baos);
		        
		        imageOutputStreamMap.put("sheet_1_image", screenshotHeader.getImage());
				
				break;
			case "SOO":
				
				options = new HTMLOptions();	
				outputStreamHTML = new ByteArrayOutputStream();
//				sheet.saveToHtml(outputStreamHTML, options);
				worksheet.saveToHtml("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOOHtml.html", options);
				
						        
				worksheet.saveToImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOOExcelToPng.png");
//				worksheet.saveToHtml(outputStreamHTML, options);
//				uploadFile("sheet2.html", outputStreamHTML);
				try {
					driver = new ChromeDriver(options1);
//					driver.get("https://seamlessserver.s3.amazonaws.com/signinlogo/fsimages/sheet2.html");
					driver.get("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOOHtml.html");
					driver.manage().window().maximize();
				} catch (Exception e) {
					System.out.println("Exception Message:-- "+e.getMessage());
					e.printStackTrace();
				}
						
				findElement = driver.findElement(By.tagName("table"));
				
				screenshotHeader = new AShot().coordsProvider(new WebDriverCoordsProvider()).shootingStrategy(ShootingStrategies.viewportPasting(100)).takeScreenshot(driver, findElement);
			    try {
			        ImageIO.write(screenshotHeader.getImage(),"png", new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOOHtmlToImage.png"));
			    } catch (IOException e) {
			        e.printStackTrace();
			    }
			    
			    pngFile = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOO.png");
				ImageIO.write(screenshotHeader.getImage(), "png", pngFile);
//		        screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
//		        outputStream = new ByteArrayOutputStream();
//		        try {
//					FileUtils.copyFile(screenshot, outputStream);
//				} catch (IOException e) {
//					e.printStackTrace();
//				}
//		        isFromFirstData = new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
//		        image = ImageIO.read(isFromFirstData);		                
//		        
//		        croppedImage = image.getSubimage(5, 50, 530, 510);
//		        
//		        pngFile = new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOO.png");
//				ImageIO.write(croppedImage, "png", pngFile);
//		        baos = new ByteArrayOutputStream();
		        
//		        ImageIO.write(croppedImage, "png", baos);
			    
		        imageOutputStreamMap.put("sheet_2_image", screenshotHeader.getImage());
				break;
				
			case "SCNA":
				
//				options = new HTMLOptions();	
//				outputStreamHTML = new ByteArrayOutputStream();
////				sheet.saveToHtml(outputStreamHTML, options);
//				worksheet.saveToHtml("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SCNA.html", options);
//				worksheet.saveToImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SCNAExcelToPng.png");
////				worksheet.saveToHtml(outputStreamHTML, options);
////				uploadFile("sheet2.html", outputStreamHTML);
//				try {
//					driver = new ChromeDriver(options1);
//					driver.get("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SCNA.html");
//					driver.manage().window().maximize();
//				} catch (Exception e) {
//					System.out.println("Exception Message:-- "+e.getMessage());
//					e.printStackTrace();
//				}
//						
//				findElement = driver.findElement(By.tagName("table"));
//				
//				screenshotHeader = new AShot().coordsProvider(new WebDriverCoordsProvider()).shootingStrategy(ShootingStrategies.viewportPasting(100)).takeScreenshot(driver, findElement);
//			    try {
//			        ImageIO.write(screenshotHeader.getImage(),"png", new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SCNAImage.png"));
//			    } catch (IOException e) {
//			        e.printStackTrace();
//			    }
//			    			    
//		        imageOutputStreamMap.put("sheet_3_image", screenshotHeader.getImage());
				break;
			case "SOI":
				
//				options = new HTMLOptions();	
//				outputStreamHTML = new ByteArrayOutputStream();
////				sheet.saveToHtml(outputStreamHTML, options);
//				worksheet.saveToHtml("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOI.html", options);
//				worksheet.saveToImage("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOIExcelToPng.png");
////				worksheet.saveToHtml(outputStreamHTML, options);
////				uploadFile("sheet2.html", outputStreamHTML);
//				try {
//					driver = new ChromeDriver(options1);
////					driver.get("https://seamlessserver.s3.amazonaws.com/signinlogo/fsimages/sheet2.html");
//					driver.get("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOI.html");
//					driver.manage().window().maximize();
//				} catch (Exception e) {
//					System.out.println("Exception Message:-- "+e.getMessage());
//					e.printStackTrace();
//				}
//						
//				findElement = driver.findElement(By.tagName("table"));
//				
//				screenshotHeader = new AShot().coordsProvider(new WebDriverCoordsProvider()).shootingStrategy(ShootingStrategies.viewportPasting(100)).takeScreenshot(driver, findElement);
//			    try {
//			        ImageIO.write(screenshotHeader.getImage(),"png", new File("C:\\Users\\RaviKumar(JAI)\\OneDrive - Formidium Corp\\Desktop\\Excel File\\SOIImage.png"));
//			    } catch (IOException e) {
//			        e.printStackTrace();
//			    }
//			    			    
//		        imageOutputStreamMap.put("sheet_4_image", screenshotHeader.getImage());
				break;
			}
		}
		
		return imageOutputStreamMap;
	}
	
	public static Map<String, InputStream> captureImageFromExcelAspose(XSSFWorkbook workbookFile) throws Exception {

		Map<String, InputStream> imageOutputStreamMap = new HashMap<String, InputStream>();
		
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		workbookFile.write(bos);
		byte[] barray = bos.toByteArray();
		InputStream is = new ByteArrayInputStream(barray);
		Workbook workbook = new Workbook(is);
		workbook.calculateFormula();
		
		Style defaultStyle = workbook.getDefaultStyle();

		defaultStyle.getFont().setName("sans-serif");
		defaultStyle.getFont().setSize(10);
		defaultStyle.setTextWrapped(true);
		defaultStyle.getFont().setCharset(0);
		workbook.setDefaultStyle(defaultStyle);
		
		workbook.getWorksheets().get(0).getPageSetup().setPrintArea("A5:A27:B5:B27");
		workbook.getWorksheets().get(0).getPageSetup().setLeftMargin(0);
		workbook.getWorksheets().get(0).getPageSetup().setRightMargin(0);
		workbook.getWorksheets().get(0).getPageSetup().setTopMargin(0);
		workbook.getWorksheets().get(0).getPageSetup().setBottomMargin(0);
		
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setOnePagePerSheet(true);
		options.setImageType(ImageType.EMF);
		options.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
		options.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
				RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
		options.setHorizontalResolution(300);
		options.setVerticalResolution(300);
		options.setTiffCompression(TiffCompression.COMPRESSION_LZW);
		options.setQuality(100);
		options.setCheckWorkbookDefaultFont(true);
		options.setDefaultFont("sans-serif");
		
		SheetRender sr = new SheetRender(workbook.getWorksheets().get(0), options);
		
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		sr.toImage(0, outStream);
		
		InputStream isFromFirstData = new ByteArrayInputStream(outStream.toByteArray()); 
		imageOutputStreamMap.put("image_sheet_1", isFromFirstData);
		
		workbook.getWorksheets().get(1).getPageSetup().setPrintArea("A5:A26:B5:B26");
		
		sr = new SheetRender(workbook.getWorksheets().get(1), options);
		
		outStream = new ByteArrayOutputStream();
		sr.toImage(0, outStream);
		
		isFromFirstData = new ByteArrayInputStream(outStream.toByteArray()); 
		imageOutputStreamMap.put("image_sheet_2", isFromFirstData);
		
		return imageOutputStreamMap;
	}
	
}