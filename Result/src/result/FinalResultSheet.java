package result;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class FinalResultSheet {

	List<String> unMergedHeaderCell;
	int totalMarks;

	public FinalResultSheet(List<String> unMergedHeaderCell) {
		this.unMergedHeaderCell = unMergedHeaderCell;
		totalMarks = unMergedHeaderCell.size() * 100;
		unMergedHeaderCell.add("Total");
		unMergedHeaderCell.add("%age");
	}

	public void createFinalResultSheets(Sheet sheet, List<Student> students,
			Map<String, String> inputValues) throws Exception {
		Row row;
		int colNum;
		int rowNum = 0;
		int studentTotal;

		rowNum = createFinalResultHeader(sheet, inputValues, rowNum);
		for (Student student : students) {
			studentTotal = 0;
			row = sheet.createRow(rowNum++);
			colNum = 0;
			createCell(ExcelUtils.borderCellStyle, row, student.getRollNum(),
					colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, student.getSrn(),
					colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, student.getName(),
					colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row,
					student.getFathersName(), colNum++, false);
			for (Subject subject : student.getSubjects()) {
				int index = unMergedHeaderCell.indexOf(subject.getName());
				createCell(ExcelUtils.borderCellStyle, row,
						subject.getSubjectTotal(), colNum + index, false);
				studentTotal = studentTotal + subject.getSubjectTotal();
			}
			colNum = colNum + unMergedHeaderCell.size() - 2;
			createCell(ExcelUtils.borderCellStyle, row, studentTotal, colNum++, false);
			
			createCell(ExcelUtils.borderCellStyle, row, String.format("%.2f",((float)(studentTotal*100))/totalMarks), colNum++, false);
			createCell(ExcelUtils.borderCellStyle, row, "", colNum++, false);
			colNum = 0;
		}
	}

	public int createFinalResultHeader(Sheet sheet,
			Map<String, String> inputValues, int rowNum) throws Exception {

		int colNum = 0;
		rowNum = 1;
		String[] mergedHeaderCell = new String[] { "RollNo", "SRN",
				"Student's Name", "Father's Name" };

		Row headerRow = sheet.createRow(rowNum++);
		Row secondHeaderRow = sheet.createRow(rowNum++);
		Row firstRow = sheet.createRow(rowNum++);
		Row secondRow = sheet.createRow(rowNum++);
		ExcelUtils.createSingleMergedCell(ExcelUtils.headerStyle, headerRow,
				inputValues.get("schoolHeader"), 0, 11);

		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle,
				secondHeaderRow, "Annual Result  " + inputValues.get("session")
						+ "               Class: " + inputValues.get("class"),
				4, 4);

		for (String label : mergedHeaderCell) {
			createCell(ExcelUtils.basicRowStyle, firstRow, label, colNum++,
					true);
		}
		for (String label : unMergedHeaderCell) {
			createCell(ExcelUtils.basicRowStyle, firstRow, label, colNum,
					false);
			if (label.equals("Total"))
				createCell(ExcelUtils.basicRowStyle, secondRow, totalMarks,
						colNum++, false);
			else
				createCell(ExcelUtils.basicRowStyle, secondRow, 100,
						colNum++, false);
		}
		createCell(ExcelUtils.basicRowStyle, firstRow, "Remarks", colNum++,
				true);
		return rowNum;
	}

	private static int mergeRegions(Sheet sheet, int rowNum, int colNum) {
		sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + 1, colNum,
				colNum));

		return rowNum + 1;
	}

	private static void createCell(CellStyle style, Row row, String label,
			int colNum, boolean isMerged) {
		Cell cell = row.createCell(colNum);

		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue(label);
		cell.setCellStyle(style);

		if (isMerged) {
			int rowNum = mergeRegions(row.getSheet(), row.getRowNum(), colNum);
			Row otherRow = row.getSheet().getRow(rowNum) == null ? row
					.getSheet().createRow(rowNum) : row.getSheet().getRow(
					rowNum);
			otherRow.createCell(colNum).setCellStyle(style);
		}
		row.getSheet().autoSizeColumn(colNum, true);
	}

	private static void createCell(CellStyle style, Row row, Integer label,
			int colNum, boolean isMerged) {
		Cell cell = row.createCell(colNum);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(label);

		cell.setCellStyle(style);
		if (isMerged)
			mergeRegions(row.getSheet(), row.getRowNum(), colNum);
		row.getSheet().autoSizeColumn(colNum, true);
	}

}
