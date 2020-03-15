package result;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class Reader {
	public List<Student> readFile(final String fileName) throws Exception {
		final FileInputStream input = new FileInputStream(fileName);
		final Workbook workbook = new HSSFWorkbook(input);
		Sheet sheet = null;
		final List<Student> students = new ArrayList<Student>();
		final int sheetsCount = workbook.getNumberOfSheets();
		try {
			for (int index = 0; index < sheetsCount; ++index) {
				sheet = workbook.getSheetAt(index);
				final int rowStart = 7;
				for (int rowEnd = sheet.getPhysicalNumberOfRows() - 1, i = rowStart; i < rowEnd; ++i) {
					final Row row = sheet.getRow(i);
					if (row != null) {
						this.createStudent(row, students);
					}
				}
			}
		} finally {
			if (input != null) {
				input.close();
			}
		}
		if (input != null) {
			input.close();
		}
		return students;
	}

	private void createStudent(final Row row, final List<Student> students) throws Exception {
		Cell cell = null;
		Student student = new Student();
		String value = "";
		List<Integer> marks = new ArrayList<Integer>();
		final String subjectName = row.getSheet().getSheetName();
		try {
			for (int i = 0; i < row.getPhysicalNumberOfCells(); ++i) {
				cell = row.getCell(i);
				if (cell != null) {
					cell.setCellType(1);
					value = cell.getStringCellValue();
					if (value == null) {
						value = "";
					}
					switch (i) {
					case 0: {
						student.setRollNum(Integer.valueOf(Integer.parseInt(value)));
						if (students.contains(student)) {
							student = students.get(students.indexOf(student));
							break;
						}
						break;
					}
					case 1: {
						student.setSrn(value);
						break;
					}
					case 2: {
						student.setName(value);
						break;
					}
					case 3: {
						student.setFathersName(value);
						break;
					}
					default: {
						marks.add(Integer.parseInt(cell.getStringCellValue()));
						break;
					}
					}
				}
			}
			final Subject subject = new Subject(subjectName, marks);
			subject.setSubjectTotal(Integer.valueOf(marks.get(marks.size() - 1)));
			student.addSubjects(subject);
			if (!students.contains(student)) {
				students.add(student);
			}
			marks = new ArrayList<Integer>();
		} catch (Exception e) {
			throw new Exception(student.getRollNum() + "  " + student.getName() + " has error in sheet : " + subjectName
					+ " Please Correct the marks details.");
		}
	}

	public List<String> getSheetNames(final String fileName) throws Exception {
		final FileInputStream input = new FileInputStream(fileName);
		final Workbook workbook = new HSSFWorkbook(input);
		final List<String> sheetNames = new ArrayList<String>();
		for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
			sheetNames.add(workbook.getSheetName(i));
		}
		return sheetNames;
	}
}