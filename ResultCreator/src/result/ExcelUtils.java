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

	public ExcelUtils(final Workbook workbook) {
		borderStyle(workbook);
		headerRowStyle(workbook);
		basicRowStyle(workbook, ExcelUtils.borderCellStyle);
		basicFormulaRowStyle(workbook, ExcelUtils.basicRowStyle);
		boldRowStyle(workbook);
	}

	public static CellStyle borderStyle(final Workbook workbook) {
		ExcelUtils.borderCellStyle = workbook.createCellStyle();
		final Font font = workbook.createFont();
		font.setBoldweight((short) 700);
		ExcelUtils.borderCellStyle.setBorderBottom((short) 1);
		ExcelUtils.borderCellStyle.setBorderLeft((short) 1);
		ExcelUtils.borderCellStyle.setBorderRight((short) 1);
		ExcelUtils.borderCellStyle.setBorderTop((short) 1);
		ExcelUtils.borderCellStyle.setVerticalAlignment((short) 0);
		ExcelUtils.borderCellStyle.setAlignment((short) 2);
		ExcelUtils.borderCellStyle.setWrapText(true);
		ExcelUtils.borderCellStyle.setFont(font);
		return ExcelUtils.borderCellStyle;
	}

	public static CellStyle basicRowStyle(final Workbook workbook, final CellStyle cloningStyle) {
		(ExcelUtils.basicRowStyle = workbook.createCellStyle()).cloneStyleFrom(cloningStyle);
		final HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
		final short colorIndex = 45;
		palette.setColorAtIndex(colorIndex, (byte) (-18), (byte) (-20), (byte) (-31));
		ExcelUtils.basicRowStyle.setFillForegroundColor(colorIndex);
		ExcelUtils.basicRowStyle.setFillBackgroundColor(colorIndex);
		ExcelUtils.basicRowStyle.setFillPattern((short) 1);
		ExcelUtils.basicRowStyle.setWrapText(true);
		return ExcelUtils.basicRowStyle;
	}

	public static CellStyle basicFormulaRowStyle(final Workbook workbook, final CellStyle cloningStyle) {
		(ExcelUtils.basicFormulaRowStyle = workbook.createCellStyle()).cloneStyleFrom(cloningStyle);
		ExcelUtils.basicFormulaRowStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
		return ExcelUtils.basicFormulaRowStyle;
	}

	public static CellStyle headerRowStyle(final Workbook workbook) {
		ExcelUtils.headerStyle = workbook.createCellStyle();
		final Font font = workbook.createFont();
		font.setBoldweight((short) 700);
		font.setFontHeightInPoints((short) 18);
		ExcelUtils.headerStyle.setFont(font);
		ExcelUtils.headerStyle.setAlignment((short) 2);
		return ExcelUtils.headerStyle;
	}

	public static CellStyle boldRowStyle(final Workbook workbook) {
		ExcelUtils.boldStyle = workbook.createCellStyle();
		final Font font = workbook.createFont();
		font.setBoldweight((short) 700);
		font.setFontHeightInPoints((short) 12);
		ExcelUtils.boldStyle.setFont(font);
		ExcelUtils.boldStyle.setAlignment((short) 1);
		ExcelUtils.boldStyle.setWrapText(true);
		return ExcelUtils.boldStyle;
	}

	public static int createSingleMergedCell(final CellStyle style, final Row row, final String label, final int colNum,
			final int colSpan) {
		for (int i = 0; i <= colSpan; ++i) {
			final Cell cell = row.createCell(colNum + i);
			cell.setCellType(1);
			cell.setCellStyle(style);
		}
		row.getCell(colNum).setCellValue(label);
		row.getSheet()
				.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), colNum, colNum + colSpan));
		return colNum + colSpan + 1;
	}
}