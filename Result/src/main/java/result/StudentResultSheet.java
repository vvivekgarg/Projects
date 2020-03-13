package result;

import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class StudentResultSheet {

	public StudentResultSheet(Workbook workbook) {
	}

	public void createStudentResultSheets(Sheet sheet, List<Student> students,
			Map<String, String> inputValues) throws Exception {
		Row row;
		int colNum=0;
		String formula;
		CellReference cr1;
		CellReference cr2;
		int rowNum = 0;
		int startRow;
		int endRow;
		int totalColumns;
		for (Student student : students) {
			rowNum = createStudentResultHeader(sheet, student, inputValues,
					rowNum);
			 startRow=rowNum;
			for (Subject subject : student.getSubjects()) {
				row = sheet.createRow(rowNum++);
				colNum = 0;
				createCell(ExcelUtils.borderCellStyle, row, subject.getName(),
						colNum++, false);
				for (Integer marks : subject.getMarks()) {
					createCell(ExcelUtils.borderCellStyle, row, marks,
							colNum++, false);

				}
			}
			 endRow=rowNum-1;
			 totalColumns=colNum-1;
			row = sheet.createRow(rowNum++);
			colNum = 0;
			createCell(ExcelUtils.borderCellStyle, row,"Total",
					colNum++, false);
			for (int i = 1; i <= totalColumns; i++) {
				cr1= new CellReference(startRow, i);
				cr2= new CellReference(endRow, i);
				formula= "SUM("+cr1.formatAsString()+":"+cr2.formatAsString()+")";
				createFormulaCell(ExcelUtils.basicRowStyle, row, formula, i, false);
			}
			
			row = sheet.createRow(rowNum++);
			createCell(ExcelUtils.basicRowStyle, row,"Overall",
					18, false);
			cr1= new CellReference(rowNum-2, totalColumns);
			formula=cr1.formatAsString()+"/"+"7";
			createFormulaCell(ExcelUtils.basicFormulaRowStyle, row, formula, 19, false);
			rowNum++;
		}

	}

	public int createStudentResultHeader(Sheet sheet, Student student,
			Map<String, String> inputValues, int rowNum) throws Exception {

		String[] mergedHeaderCell = new String[] { "Subject", "May", "July",
				"Aug", "Oct", "Nov", "Jan", "Feb", "Total", "Dividing", "Sept",
				"Dividing", "March", "Dividing" };

		String[] unMergedHeaderCell = new String[] { "WB", "CRP", "Attendance",
				"Total", "Dividing" };

		String[] maxMarks = new String[] { "20", "20", "20", "20", "20", "20",
				"20", "140", "20*7/7=20", "40", "40/2=20", "80", "80/2=40",
				"20", "10", "10", "40", "40/2=20", "100" };

		int colNum = 0;
		Row headerRow = sheet.createRow(rowNum++);
		Row secondHeaderRow = sheet.createRow(rowNum++);
		Row basicRow = sheet.createRow(rowNum++);
		Row firstRow = sheet.createRow(rowNum++);
		Row secondRow = sheet.createRow(rowNum++);
		Row marksRow = sheet.createRow(rowNum++);
		Row miscRow = sheet.createRow(rowNum++);

		// School Header
		ExcelUtils.createSingleMergedCell(ExcelUtils.headerStyle, headerRow,
				inputValues.get("schoolHeader"), 3, 11);

		// Result Header
		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle,
				secondHeaderRow,
				"Annual Result  " + inputValues.get("session"), 0, 4);

		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle,
				secondHeaderRow, "Class: " + inputValues.get("class"), 6, 5);

		ExcelUtils.createSingleMergedCell(ExcelUtils.boldStyle,
				secondHeaderRow, "Roll No. " + student.getRollNum(), 13, 2);

		// Student Detais row
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "Name of Student:", colNum, 2);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, student.getName(), colNum, 5);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "SRN No.", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, student.getSrn(), colNum, 2);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "Section", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "", colNum, 0);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "Remarks", colNum, 1);
		colNum = ExcelUtils.createSingleMergedCell(ExcelUtils.basicRowStyle,
				basicRow, "", colNum, 0);
		colNum = 0;

		// Month row
		for (String lable : mergedHeaderCell) {
			createCell(ExcelUtils.borderCellStyle, firstRow, lable, colNum++,
					true);
		}
		ExcelUtils.createSingleMergedCell(ExcelUtils.borderCellStyle, firstRow,
				"Continuous and comprehensive evaluation", colNum, 4);

		for (String label : unMergedHeaderCell) {
			createCell(ExcelUtils.borderCellStyle, secondRow, label, colNum++,
					false);
		}
		createCell(ExcelUtils.borderCellStyle, firstRow, "Final Result",
				colNum++, true);

		// Marks Row
		colNum = 0;
		createCell(ExcelUtils.borderCellStyle, marksRow, "", colNum++, false);
		for (String label : maxMarks) {
			createCell(ExcelUtils.borderCellStyle, marksRow, label, colNum++,
					false);
		}

		createCell(ExcelUtils.basicRowStyle, miscRow, "A", 9, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "B", 11, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "C", 13, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "D", 18, false);
		createCell(ExcelUtils.basicRowStyle, miscRow, "A+B+C+D", 19, false);
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
		cell.setCellValue(label);
		cell.setCellType(Cell.CELL_TYPE_STRING);
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
		cell.setCellValue(label);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellStyle(style);

		row.getSheet().autoSizeColumn(colNum, true);
	}

	private static void createFormulaCell(CellStyle style, Row row, String label,
			int colNum, boolean isMerged) {
		Cell cell = row.createCell(colNum);
		cell.setCellType(Cell.CELL_TYPE_FORMULA);
		cell.setCellFormula(label);
		cell.setCellStyle(style);
		row.getSheet().autoSizeColumn(colNum, true);
	}
}
