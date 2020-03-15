package result;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Map;

public class StudentResultSheet {
	public StudentResultSheet(final Workbook workbook) {
	}

	public Integer createStudentResultSheets(final Sheet sheet, final Student student,
			final Map<String, String> inputValues, Integer rowNum) throws Exception {
		int colNum = 0;
		rowNum = this.createStudentResultHeader(sheet, student, inputValues, rowNum);
		final int startRow = rowNum;
		for (final Subject subject : student.getSubjects()) {
			final Integer n = rowNum;
			rowNum = n + 1;
			final Row row = sheet.createRow(n);
			colNum = 0;
			createCell(ExcelUtils.borderCellStyle, row, subject.getName(), colNum++, false);
			for (final Integer marks : subject.getMarks()) {
				createCell(ExcelUtils.borderCellStyle, row, marks, colNum++, false);
			}
		}
		final int endRow = rowNum - 1;
		final int totalColumns = colNum - 1;
		final Integer n2 = rowNum;
		rowNum = n2 + 1;
		Row row = sheet.createRow(n2);
		colNum = 0;
		createCell(ExcelUtils.borderCellStyle, row, "Total", colNum++, false);
		for (int i = 1; i <= totalColumns; ++i) {
			final CellReference cr1 = new CellReference(startRow, i);
			final CellReference cr2 = new CellReference(endRow, i);
			final String formula = "SUM(" + cr1.formatAsString() + ":" + cr2.formatAsString() + ")";
			createFormulaCell(ExcelUtils.basicRowStyle, row, formula, i, false);
		}
		final Integer n3 = rowNum;
		rowNum = n3 + 1;
		row = sheet.createRow(n3);
		++rowNum;
		return rowNum;
	}

	public int createStudentResultHeader(final Sheet sheet, final Student student,
			final Map<String, String> inputValues, int rowNum) throws Exception {
		final String[] mergedHeaderCell = { "Subject", "1st M.T.", "2nd M.T.", "3rd M.T.", "4th M.T.", "5th M.T.",
				"6th M.T.", "7th M.T.", "Total", "Dividing", "H.Y. Sept", "Dividing", "March", "Practical Sc.",
				"Practical Oth." };
		final String[] unMergedHeaderCell = { "WB", "CRP", "Attendance", "Total", "Dividing" };
		final String[] maxMarks = { "24", "24", "24", "24", "24", "0", "0", "120", "120/10=12", "40", "40/10=4", "80",
				"", "", "20", "10", "10", "40", "40/10=4", "100" };
		int colNum = 0;
		final Row headerRow = sheet.createRow(rowNum++);
		final Row secondHeaderRow = sheet.createRow(rowNum++);
		final Row basicRow = sheet.createRow(rowNum++);
		final Row firstRow = sheet.createRow(rowNum++);
		final Row secondRow = sheet.createRow(rowNum++);
		final Row marksRow = sheet.createRow(rowNum++);
		final Row miscRow = sheet.createRow(rowNum++);
		ExcelUtils.createSingleMergedCell(ExcelUtils.headerStyle, headerRow, inputValues.get("schoolHeader"), 3, 11);
		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle, secondHeaderRow,
				"Annual Result  " + inputValues.get("session"), 0, 4);
		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle, secondHeaderRow, "Class: " + inputValues.get("class"),
				6, 5);
		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle, secondHeaderRow, "Roll No. " + student.getRollNum(), 13,
				2);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "Name of Student:", colNum, 2);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, student.getName(), colNum, 5);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "SRN No.", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, student.getSrn(), colNum, 2);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "Section", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "", colNum, 0);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "Remarks", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle, basicRow, "", colNum, 1);
		colNum = 0;
		String[] array;
		for (int length = (array = mergedHeaderCell).length, i = 0; i < length; ++i) {
			final String lable = array[i];
			createCell(ExcelUtils.borderCellStyle, firstRow, lable, colNum++, true);
		}
		ExcelUtils.createSingleMergedCell(ExcelUtils.borderCellStyle, firstRow,
				"Continuous and comprehensive evaluation", colNum, 4);
		String[] array2;
		for (int length2 = (array2 = unMergedHeaderCell).length, j = 0; j < length2; ++j) {
			final String label = array2[j];
			createCell(ExcelUtils.borderCellStyle, secondRow, label, colNum++, false);
		}
		createCell(ExcelUtils.borderCellStyle, firstRow, "Final Result", colNum++, true);
		colNum = 0;
		createCell(ExcelUtils.borderCellStyle, marksRow, "", colNum++, false);
		String[] array3;
		for (int length3 = (array3 = maxMarks).length, k = 0; k < length3; ++k) {
			final String label = array3[k];
			createCell(ExcelUtils.borderCellStyle, marksRow, label, colNum++, false);
		}
		createCell(ExcelUtils.basicRowStyle, miscRow, "A", 9, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "B", 11, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "C", 12, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "D", 13, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "E", 14, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "F", 19, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "A+B+C+D+E+F", 20, false);
		return rowNum;
	}

	private static int mergeRegions(final Sheet sheet, final int rowNum, final int colNum) {
		sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + 1, colNum, colNum));
		return rowNum + 1;
	}

	private static void createCell(final CellStyle style, final Row row, final String label, final int colNum,
			final boolean isMerged) {
		final Cell cell = row.createCell(colNum);
		cell.setCellValue(label);
		cell.setCellType(1);
		cell.setCellStyle(style);
		if (isMerged) {
			final int rowNum = mergeRegions(row.getSheet(), row.getRowNum(), colNum);
			final Row otherRow = (row.getSheet().getRow(rowNum) == null) ? row.getSheet().createRow(rowNum)
					: row.getSheet().getRow(rowNum);
			otherRow.createCell(colNum).setCellStyle(style);
		}
	}

	private static void createCell(final CellStyle style, final Row row, final Integer label, final int colNum,
			final boolean isMerged) {
		final Cell cell = row.createCell(colNum);
		cell.setCellValue(label);
		cell.setCellType(0);
		cell.setCellStyle(style);
	}

	private static void createFormulaCell(final CellStyle style, final Row row, final String label, final int colNum,
			final boolean isMerged) {
		final Cell cell = row.createCell(colNum);
		cell.setCellType(2);
		cell.setCellFormula(label);
		cell.setCellStyle(style);
	}
}