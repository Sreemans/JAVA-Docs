// <dependency>
// 			<groupId>org.apache.poi</groupId>
// 			<artifactId>poi-ooxml</artifactId>
// 		</dependency>

import java.io.ByteArrayOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.Comparator;

import org.apache.commons.lang.BooleanUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Excel {
	public static final String YES = "Yes";
	public static final String NO = "No";

	private <T> void copyHeaderCellValue(XSSFSheet sheet, int row, int cell, T data,XSSFCellStyle cellStyle) {
		XSSFRow r = sheet.getRow(row);
		if (r == null) {
			r = sheet.createRow(row);
		}
		XSSFCell xcell = r.getCell(cell);
		if (xcell == null) {
			xcell = r.createCell(cell, =.STRING);
					}
		xcell.setCellType(CellType.STRING);
		
		xcell.setCellStyle(cellStyle);
		
		if (data != null) {
			if ((data instanceof String) || (data instanceof Double)) {

				xcell.setCellValue(replaceNull(data.toString()));
			} else {
				xcell.setCellValue(replaceNull((String) data));
			}
		} else {
			xcell.setCellValue("");
		}
	}

	private String replaceNull(String val) {
		if (val == null) {
			return "";
		}
		return val;
	}

	private String getYesNoString(Boolean val) {
		if (val != null) {
			return val ? YES : NO;
		}
		return "";
	}

	private <T> void copyCellValue(XSSFSheet sheet, int row, int cell, T data, boolean blankNoBoolean,
			XSSFCellStyle cellStyle) {
		XSSFRow r = sheet.getRow(row);
		if (r == null) {
			r = sheet.createRow(row);
		}
		XSSFCell xcell = r.getCell(cell);
		if (xcell == null) {
			xcell = r.createCell(cell, CellType.STRING);
		}
		xcell.setCellType(CellType.STRING);
		if (data != null) {
			if ((data instanceof String) || (data instanceof Double)) {

				xcell.setCellValue(replaceNull(data.toString()));
			} else if (data instanceof Boolean) {
				if (blankNoBoolean)
					xcell.setCellValue((Boolean) data ? YES : "");
				else
					xcell.setCellValue(getYesNoString((Boolean) data));
			} else if (data instanceof Date) {
				SimpleDateFormat sdf = new SimpleDateFormat("MMMMM dd yyyy");
				xcell.setCellValue(sdf.format((Date) data));
			} else {
				xcell.setCellValue(replaceNull((String) data));
			}
		} else {
			xcell.setCellValue("");
		}
		if (cellStyle != null) {
			xcell.setCellStyle(cellStyle);
		}
	}

	private List<String> emptyListCols(int colNums) {
		List<String> emptyList = new ArrayList<>();
		for (int i = 0; i < colNums; i++) {
			emptyList.add("");
		}

		return emptyList;
	}

	private void createEmptyRow(XSSFWorkbook workbook, XSSFSheet sheet, int cols) {

		List<String> emptyList = emptyListCols(cols);
		createRow(workbook, sheet, false, false, emptyList);
	}

	private List<String> GetResultColumns() {
		List<String> columns = new ArrayList<String>();
		columns.add(("lp_result"));
		columns.add(("lp_result_desc"));
		columns.add(("lp_result_desc_fr"));
		columns.add(("lp_method"));
		columns.add(("lp_method_desc"));
		columns.add(("lp_method_desc_fr"));
		columns.add(("lp_status"));
		columns.add(("lp_priority"));
		columns.add(("lp_alt_result"));

		return columns;
	}

	private CellStyle getHeaderRowStyle(XSSFWorkbook workbook) {

		XSSFCellStyle msStyle = workbook.createCellStyle();
		XSSFFont msFont = workbook.createFont();
		msFont.setFontName("Calibri");
		msFont.setBold(true);
		msFont.setItalic(false);
		msStyle.setFont(msFont);
		msStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		msStyle.setAlignment(HorizontalAlignment.LEFT);
		msStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		msStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		msStyle.setBorderBottom(BorderStyle.THIN);
		msStyle.setBorderLeft(BorderStyle.THIN);
		msStyle.setBorderTop(BorderStyle.THIN);
		msStyle.setBorderRight(BorderStyle.THIN);

		return msStyle;
	}

	private CellStyle getSubHeaderRowStyle(XSSFWorkbook workbook) {

		XSSFCellStyle msSubStyle = workbook.createCellStyle();
		XSSFFont msFont = workbook.createFont();
		msFont.setFontName("Calibri");
		msFont.setBold(true);
		msFont.setItalic(false);
		msSubStyle.setFont(msFont);
		msSubStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		msSubStyle.setAlignment(HorizontalAlignment.LEFT);
		msSubStyle.setBorderBottom(BorderStyle.THIN);
		msSubStyle.setBorderLeft(BorderStyle.THIN);
		msSubStyle.setBorderTop(BorderStyle.THIN);
		msSubStyle.setBorderRight(BorderStyle.THIN);
		msSubStyle.setWrapText(true);

		return msSubStyle;
	}
	private CellStyle getRowStyle(XSSFWorkbook workbook) {
		
		XSSFCellStyle msRowStyle = workbook.createCellStyle();
		msRowStyle.setBorderBottom(BorderStyle.THIN);
		msRowStyle.setBorderLeft(BorderStyle.THIN);
		msRowStyle.setBorderTop(BorderStyle.THIN);
		msRowStyle.setBorderRight(BorderStyle.THIN);
		
		return msRowStyle;
	}
	private void createRow(XSSFWorkbook workbook, XSSFSheet sheet, boolean headerRow, boolean subHeader,
			List<String> listData) {

		XSSFRow clientRow = sheet.createRow(0);
		for (int i = 0; i < listData.size(); i++) {

			XSSFCell clientCell = clientRow.createCell(i);
			if (!StringUtils.isEmpty(listData.get(i))) {
				clientCell.setCellValue(listData.get(i));
			}
			if (headerRow) {
				clientCell.setCellStyle(getHeaderRowStyle(workbook));
			} else if (subHeader) {
				clientCell.setCellStyle(getSubHeaderRowStyle(workbook));
			} else {
				clientCell.setCellStyle(getRowStyle(workbook));
			}

		}

	}

	public byte[] generateExcelReportStream(String companyOid, String projectOid) throws Exception {

		byte[] output = null;
		XSSFWorkbook workbook = new XSSFWorkbook();

		ByteArrayOutputStream out = new ByteArrayOutputStream();
		XSSFSheet sheet = workbook.createSheet("lp_client_details");
		createEmptyRow(workbook, sheet, 0);
		XSSFCellStyle msHeaderStyle = workbook.createCellStyle();
		XSSFFont msFont = workbook.createFont();
		msFont.setFontName("Calibri");
		msFont.setBold(true);
		msFont.setItalic(false);
		msHeaderStyle.setFont(msFont);

		copyHeaderCellValue(sheet, 0, 1, "lp_pr_calculations_report", msHeaderStyle);
		sheet.autoSizeColumn(2);

		XSSFSheet resultSheet = workbook.createSheet(("lp_calculation_result"));

		List<String> resultColumns = GetResultColumns();
		createRow(workbook, resultSheet, true, false, resultColumns);
		for(int columnIndex = 0; columnIndex < 10; columnIndex++) {
			resultSheet.autoSizeColumn(columnIndex);
		}
	}
}