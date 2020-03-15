package result;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class FinalResultSheet {
	List<String> unMergedHeaderCell;
	int totalMarks;

	public FinalResultSheet(final List<String> unMergedHeaderCell) {
		this.unMergedHeaderCell = unMergedHeaderCell;
		this.totalMarks = unMergedHeaderCell.size() * 100;
		unMergedHeaderCell.add("Total");
		unMergedHeaderCell.add("%age");
	}

	public void createFinalResultSheets(final Sheet sheet, final List<Student> students,
			final Map<String, String> inputValues) throws Exception {
		int rowNum = 0;
		rowNum = this.createFinalResultHeader(sheet, inputValues, rowNum);
		for (final Student student : students) {
			int studentTotal = 0;
			final Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			createCell(ExcelUtils.borderCellStyle, row, student.getRollNum(), colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, student.getSrn(), colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, student.getName(), colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, student.getFathersName(), colNum++, false);
			for (final Subject subject : student.getSubjects()) {
				final int index = this.unMergedHeaderCell.indexOf(subject.getName());
				createCell(ExcelUtils.borderCellStyle, row, subject.getSubjectTotal(), colNum + index, false);
				studentTotal += subject.getSubjectTotal();
			}
			colNum = colNum + this.unMergedHeaderCell.size() - 2;
			createCell(ExcelUtils.borderCellStyle, row, studentTotal, colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row,
					String.format("%.2f", studentTotal * 100 / (float) this.totalMarks), colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, "", colNum++, false);
			colNum = 0;
		}
	}

	public int createFinalResultHeader(final Sheet sheet, final Map<String, String> inputValues, int rowNum)
			throws Exception {
		int colNum = 0;
		rowNum = 1;
		final String[] mergedHeaderCell = { "RollNo", "SRN", "Student's Name", "Father's Name" };
		final Row headerRow = sheet.createRow(rowNum++);
		final Row secondHeaderRow = sheet.createRow(rowNum++);
		final Row firstRow = sheet.createRow(rowNum++);
		final Row secondRow = sheet.createRow(rowNum++);
		ExcelUtils.createSingleMergedCell(ExcelUtils.headerStyle, headerRow, inputValues.get("schoolHeader"), 0, 11);
		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle, secondHeaderRow,
				"Annual Result  " + inputValues.get("session") + "               Class: " + inputValues.get("class"), 4,
				4);
		String[] array;
		for (int length = (array = mergedHeaderCell).length, i = 0; i < length; ++i) {
			final String label = array[i];
			createCell(ExcelUtils.basicRowStyle, firstRow, label, colNum++, true);
		}
		final Iterator<String> iterator = this.unMergedHeaderCell.iterator();
		while (iterator.hasNext()) {
			final String label = iterator.next();
			createCell(ExcelUtils.basicRowStyle, firstRow, label, colNum, false);
			if (label.equals("Total")) {
				createCell(ExcelUtils.basicRowStyle, secondRow, this.totalMarks, colNum++, false);
			} else {
				createCell(ExcelUtils.basicRowStyle, secondRow, 100, colNum++, false);
			}
		}
		createCell(ExcelUtils.basicRowStyle, firstRow, "Remarks", colNum++, true);
		return rowNum;
	}

	private static int mergeRegions(final Sheet sheet, final int rowNum, final int colNum) {
		sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + 1, colNum, colNum));
		return rowNum + 1;
	}

	private static void createCell(final CellStyle style, final Row row, final String label, final int colNum,
			final boolean isMerged) {
		final Cell cell = row.createCell(colNum);
		cell.setCellType(1);
		cell.setCellValue(label);
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
		cell.setCellType(0);
		cell.setCellValue(label);
		cell.setCellStyle(style);
		if (isMerged) {
			mergeRegions(row.getSheet(), row.getRowNum(), colNum);
		}
	}
}