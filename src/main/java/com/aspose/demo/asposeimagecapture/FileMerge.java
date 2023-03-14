package com.aspose.demo.asposeimagecapture;

import java.awt.Dimension;
import java.awt.RenderingHints;
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

import javax.imageio.ImageIO;

import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

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

public class FileMerge {

	public static InputStream writeImage(InputStream inputStream, XSSFWorkbook workbook) throws IOException {

		Map<String, OutputStream> imageURLMap = takeSnapshotFromExcel(workbook);
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

			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			workbookFile.write(bos);
			byte[] barray = bos.toByteArray();
			InputStream is = new ByteArrayInputStream(barray);

			Workbook workbook = new Workbook(is);
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
			
			FontConfigs.setDefaultFontName("Arial");
			workbook.calculateFormula();
			Style defaultStyle = workbook.getDefaultStyle();

			defaultStyle.getFont().setName("Arial");
			defaultStyle.getFont().setSize(10);
			defaultStyle.setTextWrapped(true);
			defaultStyle.getFont().setBold(false);
			workbook.setDefaultStyle(defaultStyle);
			
			workbook.getWorksheets().get(0).getPageSetup().setPrintArea("A5:A31:B5:B31");
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
			options.setHorizontalResolution(180);
			options.setVerticalResolution(180);
			options.setTiffCompression(TiffCompression.COMPRESSION_LZW);
			options.setQuality(100);
			options.setCheckWorkbookDefaultFont(true);
			options.setDefaultFont("Arial");
			
			SheetRender sr = new SheetRender(workbook.getWorksheets().get(0), options);
			ByteArrayOutputStream outStream = new ByteArrayOutputStream();
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

}