package result;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelUtils {

	static CellStyle borderCellStyle;
	static CellStyle basicRowStyle;
	static CellStyle basicFormulaRowStyle;
	static CellStyle headerStyle;
	static CellStyle boldStyle;

	public ExcelUtils(Workbook workbook) {
		borderStyle(workbook);
		headerRowStyle(workbook);
		basicRowStyle(workbook, borderCellStyle);
		basicFormulaRowStyle(workbook, basicRowStyle);
		boldRowStyle(workbook);
	}
	public static CellStyle borderStyle(Workbook workbook) {
			borderCellStyle = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			borderCellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			borderCellStyle.setBorderLeft(CellStyle.BORDER_THIN);
			borderCellStyle.setBorderRight(CellStyle.BORDER_THIN);
			borderCellStyle.setBorderTop(CellStyle.BORDER_THIN);
			borderCellStyle.setVerticalAlignment(CellStyle.VERTICAL_TOP);
			borderCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			borderCellStyle.setWrapText(true);
			borderCellStyle.setFont(font);
		return borderCellStyle;
	}

	public static CellStyle basicRowStyle(Workbook workbook,
			CellStyle cloningStyle) {
		
			basicRowStyle = workbook.createCellStyle();
			basicRowStyle.cloneStyleFrom(cloningStyle);
			HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
			short colorIndex = 45;
			palette.setColorAtIndex(colorIndex, (byte) 238, (byte) 236,
					(byte) 225);
			basicRowStyle.setFillForegroundColor(colorIndex);
			basicRowStyle.setFillBackgroundColor(colorIndex);
			basicRowStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			basicRowStyle.setWrapText(true);
		
		return basicRowStyle;
	}

	public static CellStyle basicFormulaRowStyle(Workbook workbook,
			CellStyle cloningStyle) {
		
		basicFormulaRowStyle = workbook.createCellStyle();
		basicFormulaRowStyle.cloneStyleFrom(cloningStyle);
			
			basicFormulaRowStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
		
		return basicFormulaRowStyle;
	}
	public static CellStyle headerRowStyle(Workbook workbook) {
		
			headerStyle = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			font.setFontHeightInPoints((short) 18);
			headerStyle.setFont(font);
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		
		return headerStyle;
	}

	public static CellStyle boldRowStyle(Workbook workbook) {
		
			boldStyle = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			font.setFontHeightInPoints((short) 12);
			boldStyle.setFont(font);
			boldStyle.setAlignment(CellStyle.ALIGN_LEFT);
			boldStyle.setWrapText(true);
		
		return boldStyle;
	}

	public static int createSingleMergedCell(CellStyle style, Row row,
			String label, int colNum, int colSpan) {
		Cell cell;

		for (int i = 0; i <= colSpan; i++) {
			cell = row.createCell(colNum + i);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellStyle(style);
		}
		row.getCell(colNum).setCellValue(label);
		row.getSheet().addMergedRegion(
				new CellRangeAddress(row.getRowNum(), row.getRowNum(), colNum,
						colNum + colSpan));
		row.getSheet().autoSizeColumn(colNum, true);
		return colNum + colSpan + 1;

	}

}
