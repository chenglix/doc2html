package com.doc2html;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import com.doc2html.bean.ExcelFontInfo;
import com.doc2html.bean.HtmlTable;
import com.doc2html.bean.HtmlTableTd;
import com.doc2html.bean.HtmlTableTr;
import com.doc2html.bean.dto.DocHtmlDto;
import com.doc2html.bean.dto.ExcelHtmlResultDto;
import com.doc2html.bean.dto.ExcelSheetDto;

public class Excel2Html implements Doc2Html {

	static Workbook workBook = null;

	public DocHtmlDto doc2Html(String filePath) throws IOException {
		ZipSecureFile.setMinInflateRatio(0.00000001);
		File excelFile = new File(filePath);
		FileInputStream fpis = new FileInputStream(excelFile);
		try {
			workBook = WorkbookFactory.create(fpis);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (workBook == null) {
			return null;
		}
		fpis.close();
		int sheetCount = workBook.getNumberOfSheets();
		List<ExcelSheetDto> sheetList = new ArrayList<>(sheetCount);
		for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
			Sheet sheet = workBook.getSheetAt(sheetIndex);

			HtmlTable table = new HtmlTable();
			for (Row row : sheet) {
				HtmlTableTr tr = new HtmlTableTr();
				int firstCellNum = row.getFirstCellNum();
				for (int i = 0; i < firstCellNum; i++) {
					HtmlTableTd emptyTd = new HtmlTableTd();
					tr.addTd(emptyTd);
				}
				for (Cell cell : row) {
					HtmlTableTd td = new HtmlTableTd();
					if (cell != null) {
						String tdText = null;
						int cellType = cell.getCellType();
						switch (cellType) {
							case Cell.CELL_TYPE_STRING:
								tdText = cell.getStringCellValue();
							break;
							case Cell.CELL_TYPE_NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
									tdText = FormartUtil.formartDate(cell.getDateCellValue(), cell.getCellStyle().getDataFormat());
								} else {
									tdText = FormartUtil.formartNumber1(cell.getNumericCellValue());
									/**
									 * switch (style) { //
									 * 单元格格式为百分比，不格式化会直接以小数输出 case 9: value =
									 * new DecimalFormat("0.00%").format(hcell
									 * .getNumericCellValue()); break; //
									 * DateUtil判断其不是日期格式，在这里也可以设置其输出类型 case 57:
									 * value = new SimpleDateFormat(
									 * " yyyy'年'MM'月' "
									 * ).format(hcell.getDateCellValue());
									 * break; default: value =
									 * String.valueOf(hcell.getNumericCellValue(
									 * )); break; }
									 */
								}

							default:
							break;
						}
						td.setText(tdText);
					} else {
						td.setText("&nbsp;");
					}

					CellStyle cellStyle = cell.getCellStyle();
					String backgroudColor = getCellBackground(cellStyle);
					if (backgroudColor != null) {
						td.setBackground(backgroudColor);
					}

					int height = (int) row.getHeight() / 20;
					td.setHeight(height);

					int width = (int) sheet.getColumnWidthInPixels(cell.getColumnIndex());
					td.setWidth(width);

					String align = getAlign(cellStyle);
					if (align != null) {
						td.setAlign(align);
					}

					String valign = getValigin(cellStyle);
					if (valign != null) {
						td.setValign(valign);
					}

					ExcelFontInfo font = getFont(cellStyle);
					if (font != null) {
						td.setFont(font);
					}

					tr.addTd(td);
				}
				table.addTr(tr);
			}

			List<CellRangeAddress> addresss = sheet.getMergedRegions();
			for (CellRangeAddress address : addresss) {
				int startColumn = address.getFirstColumn();
				int endColumn = address.getLastColumn();
				int startRow = address.getFirstRow();
				int endRow = address.getLastRow();

				if (startRow >= table.getTrList().size()) {
					continue;
				}
				HtmlTableTr tr = table.getTrList().get(startRow);
				HtmlTableTd td = tr.getTd(startColumn);

				int rowspan = endRow - startRow;
				int colspan = endColumn - startColumn;
				if (rowspan > 0) {
					int height = 0;
					for (int i = startRow; i <= endRow; i++) {
						height += table.getTrList().get(i).getTd(startColumn).getHeight();
					}
					td.setHeight(height);
				}
				if (colspan > 0) {
					int width = 0;
					List<HtmlTableTd> tdList = tr.getTdList();
					for (int i = startColumn; i <= endColumn; i++) {
						if (i < tdList.size()) {
							width += tdList.get(i).getWidth();
						}
					}
					td.setWidth(width);
				}
			}

			// List<ExcelPictureData> imageList = getPicture(sheet);
			// if (imageList != null && !imageList.isEmpty()) {
			// List<HtmlImage> htmlImageList = new ArrayList<HtmlImage>();
			// for (ExcelPictureData picture : imageList) {
			// HtmlImage imageItem = new HtmlImage();
			// imageItem.setPath(picture.getPath());
			//
			// int col1 = picture.getCol1();
			// int col2 = picture.getCol2();
			// int row1 = picture.getRow1();
			// int row2 = picture.getRow2();
			//
			// int width = picture.getCol2X();
			// for (int i = col1; i < col2; i++) {
			// width += table.getTr(row1).getTd(i).getWidth();
			// }
			// width -= picture.getCol1X();
			//
			// System.out.println(picture.getPath() + "：" + picture.getRow2Y());
			// int height = picture.getRow2Y();
			// for (int i = row1; i < row2; i++) {
			// height += table.getTr(i).getTd(col1).getHeight();
			// }
			// height -= picture.getRow1Y();
			//
			// int top = picture.getRow1Y();
			// for (int i = 0; i < row1; i++) {
			// top += table.getTr(i).getTd(col1).getHeight();
			// }
			// int left = picture.getCol1X();
			// for (int i = 0; i < col1; i++) {
			// left += table.getTr(row1).getTd(i).getWidth();
			// }
			//
			// imageItem.setWidth(width);
			// imageItem.setHeight(height);
			// imageItem.setTop(top);
			// imageItem.setLeft(left);
			// htmlImageList.add(imageItem);
			// }
			// table.setImageList(htmlImageList);
			// }
			for (CellRangeAddress address : addresss) {
				int startColumn = address.getFirstColumn();
				int endColumn = address.getLastColumn();
				int startRow = address.getFirstRow();
				int endRow = address.getLastRow();

				if (startRow >= table.getTrList().size()) {
					continue;
				}
				HtmlTableTr tr = table.getTrList().get(startRow);
				HtmlTableTd td = tr.getTd(startColumn);

				int rowspan = endRow - startRow;
				int colspan = endColumn - startColumn;
				if (rowspan > 0) {
					td.setRowspan(rowspan + 1);
					for (int i = startRow + 1; i <= endRow; i++) {
						for (int j = startColumn; j <= endColumn; j++) {
							table.getTrList().get(i).getTdList().set(j, null);
						}
					}
				}
				if (colspan > 0) {
					td.setColspan(colspan + 1);
					List<HtmlTableTd> tdList = tr.getTdList();
					for (int i = startColumn + 1; i <= endColumn; i++) {
						if (i < tdList.size()) {
							tdList.set(i, null);
						}
					}
				}
			}

			String html = table.toString();
			workBook.close();

			ExcelSheetDto sheetDto = new ExcelSheetDto();
			sheetDto.setSheetName(sheet.getSheetName());
			sheetDto.setSheetIndex(sheetIndex);
			sheetDto.setHtml(html);
			sheetList.add(sheetDto);
		}

		ExcelHtmlResultDto resultDto = new ExcelHtmlResultDto();
		resultDto.setSheetList(sheetList);
		resultDto.setActiveSheetIndex(workBook.getActiveSheetIndex());
		return resultDto;
	}

	private static String getValigin(CellStyle cellStyle) {
		String result = null;
		int align = cellStyle.getVerticalAlignment();
		switch (align) {
			case CellStyle.VERTICAL_CENTER:
				result = "middle";
			break;
			case CellStyle.VERTICAL_BOTTOM:
				result = "bottom";
			break;
			case CellStyle.VERTICAL_TOP:
				result = "top";
			break;
			default:
			break;
		}
		return result;
	}

	private static String getCellBackground(CellStyle cellStyle) {
		String result = null;
		Color color = cellStyle.getFillForegroundColorColor();
		if (color != null) {
			if (color instanceof XSSFColor) {
				XSSFColor xSSFColor = (XSSFColor) color;
				result = xSSFColor.getARGBHex();
				if (result != null) {
					result = "#" + result.substring(2);
				}
			} else if (color instanceof HSSFColor) {
				HSSFColor hSSFColor = (HSSFColor) color;
				result = hSSFColor.getHexString();
				if (result != null) {
					result = "#" + result.substring(2);
				}
			}
		}
		return result;
	}

	private static ExcelFontInfo getFont(CellStyle cellStyle) {
		ExcelFontInfo result = null;
		if (cellStyle instanceof XSSFCellStyle) {
			result = new ExcelFontInfo();
			XSSFCellStyle xStyle = (XSSFCellStyle) cellStyle;
			XSSFFont font = xStyle.getFont();
			if (font != null) {
				short fontSize = font.getFontHeight();
				result.setFontSize((int) fontSize / 15);
				result.setColor(getFontColor(cellStyle));
				result.setFontFamily(font.getFontName());
			}
		} else if (cellStyle instanceof HSSFCellStyle) {
			result = new ExcelFontInfo();
			HSSFCellStyle xStyle = (HSSFCellStyle) cellStyle;
			HSSFFont font = xStyle.getFont(workBook);
			if (font != null) {
				short fontSize = font.getFontHeight();
				result.setFontSize((int) fontSize / 15);
				result.setColor(getFontColor(cellStyle));
				result.setFontFamily(font.getFontName());
			}
		}
		return result;
	}

	private static String getAlign(CellStyle cellStyle) {

		String result = null;
		int align = cellStyle.getAlignment();
		switch (align) {
			case CellStyle.ALIGN_CENTER:
				result = "center";
			break;
			case CellStyle.ALIGN_LEFT:
				result = "left";
			break;
			case CellStyle.ALIGN_RIGHT:
				result = "right";
			break;
			default:
			break;
		}
		return result;
	}

	private static String getFontColor(CellStyle cellStyle) {
		String result = null;
		if (cellStyle instanceof HSSFCellStyle) {
			HSSFFont font = ((HSSFCellStyle) cellStyle).getFont(workBook);
			if (font != null) {
				HSSFColor color = font.getHSSFColor((HSSFWorkbook) workBook);
				if (color != null) {
					result = color.getHexString();
					if (result != null) {
						result = "#" + result.substring(2);
					}
				}
			}
		} else if (cellStyle instanceof XSSFCellStyle) {
			XSSFFont font = ((XSSFCellStyle) cellStyle).getFont();
			if (font != null) {
				XSSFColor color = font.getXSSFColor();
				if (color != null) {
					result = color.getARGBHex();
					if (result != null) {
						result = "#" + result.substring(2);
					}
				}
			}
		}
		return result;
	}
	//
	// private static List<ExcelPictureData> getPicture(Sheet sheet) throws
	// IOException {
	//
	// List<ExcelPictureData> resultList = null;
	// if (sheet instanceof HSSFSheet) {
	// List<HSSFPictureData> pictures = ((HSSFWorkbook)
	// workBook).getAllPictures();
	// HSSFPatriarch hp = (HSSFPatriarch) sheet.getDrawingPatriarch();
	// if (hp != null) {
	// resultList = new ArrayList<ExcelPictureData>();
	// long time = System.currentTimeMillis();
	// List<HSSFShape> shapes = hp.getChildren();
	// for (HSSFShape shape : shapes) {
	// HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
	// if (shape instanceof HSSFPicture) {
	// HSSFPicture pic = (HSSFPicture) shape;
	// int pictureIndex = pic.getPictureIndex() - 1;
	// ExcelPictureData imageItem = new ExcelPictureData();
	// imageItem.setCol1(anchor.getCol1());
	// imageItem.setCol2(anchor.getCol2());
	// imageItem.setRow1(anchor.getRow1());
	// imageItem.setRow2(anchor.getRow2());
	//
	// int rate = 3;
	// imageItem.setCol1X(anchor.getDx1() / rate);
	// imageItem.setCol2X(anchor.getDx2() / rate);
	// imageItem.setRow1Y(anchor.getDy1() / rate);
	// imageItem.setRow2Y(anchor.getDy2() / rate);
	//
	// String saveImagePath = savePicture("D:/doc2htmltest/image/excel/" + time,
	// pictures.get(pictureIndex), pictureIndex + ".jpg");
	// imageItem.setPath(saveImagePath);
	//
	// resultList.add(imageItem);
	// }
	// }
	// }
	// }
	// return resultList;
	// }

	// private static String savePicture(String folder, HSSFPictureData data,
	// String fileName) throws IOException {
	// File folderFile = new File(folder);
	// if (!folderFile.exists()) {
	// folderFile.mkdirs();
	// }
	// StringBuilder filePath = new StringBuilder();
	// filePath.append(folder).append("/").append(fileName);
	// OutputStream outp = new FileOutputStream(filePath.toString());
	// byte[] bytes = data.getData();
	// outp.write(bytes);
	// outp.close();
	// return filePath.toString();
	// }
	//
	static class FormartUtil {
		static final DecimalFormat fmt1 = new DecimalFormat("#");
		static final SimpleDateFormat dateFmt1 = new SimpleDateFormat("yyyy年M月d日");
		static final SimpleDateFormat dateFmt2 = new SimpleDateFormat("yyyy/MM/dd");
		static final SimpleDateFormat dateFmt3 = new SimpleDateFormat("yyyy/MM/dd HH:mm");
		static final SimpleDateFormat dateFmt4 = new SimpleDateFormat("yyyy/MM/dd HH:mm a");
		static final SimpleDateFormat dateFmt5 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

		protected static String formartNumber1(Number number) {
			return fmt1.format(number);
		}

		protected static String formartDate(Date date, int dateType) {
			String result = null;
			switch (dateType) {
				case 178:
					result = dateFmt1.format(date);
				break;
				case 14:
					result = dateFmt2.format(date);
				break;
				case 179:
					result = dateFmt3.format(date);
				break;
				case 181:
					result = dateFmt4.format(date);
				break;
				case 22:
					result = dateFmt5.format(date);
				break;
				default:
				break;
			}
			return result;
		}
	}

}
