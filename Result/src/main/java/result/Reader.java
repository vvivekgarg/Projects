package result;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Reader {

	public List<Student> readFile(String fileName) throws Exception {

		FileInputStream input = new FileInputStream(fileName);

		Workbook workbook = new HSSFWorkbook(input);
		Sheet sheet = null;
		List<Student> students = new ArrayList<Student>();
		int sheetsCount = workbook.getNumberOfSheets();
		try{for (int index = 0; index < sheetsCount; index++) {
			sheet = workbook.getSheetAt(index);
			int rowStart = 7;
			int rowEnd = sheet.getPhysicalNumberOfRows() - 1;
			for (int i = rowStart; i < rowEnd; i++) {
				Row row = sheet.getRow(i);
				if (row != null) {
					createStudent(row, students);

				}

			}
		}
		}finally{
			if(input!=null)
				input.close();
		}
		return students;
	}

	private void createStudent(Row row, List<Student> students) {
		Cell cell = null;
		Student student = new Student();
		String value = "";
		List<Integer> marks = new ArrayList<Integer>();

		for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
			cell = row.getCell(i);
			if (cell != null) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				value = cell.getStringCellValue();
				if (value == null)
					value = "";
				switch (i) {
				case 0:
					student.setRollNum(Integer.parseInt(value));
					if (students.contains(student))
						student = students.get(students.indexOf(student));
					break;

				case 1:
					student.setSrn(value);
					break;

				case 2:
					student.setName(value);
					break;
					
				case 3:
					student.setFathersName(value);
					break;
					
				default:
					marks.add(Integer.parseInt(cell.getStringCellValue()));
					break;
				}

			}

		}
		Subject subject = new Subject(row.getSheet().getSheetName(), marks);
		subject.setSubjectTotal(marks.get(marks.size()-1));
		student.addSubjects(subject);
		if(!students.contains(student))
		students.add(student);
		marks = new ArrayList<Integer>();
	}
	
	public List<String> getSheetNames(String fileName) throws Exception{
		FileInputStream input = new FileInputStream(fileName);
		Workbook workbook = new HSSFWorkbook(input);
		List<String> sheetNames =new ArrayList<String>();
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheetNames.add(workbook.getSheetName(i));
		}
		return sheetNames;
	}
}
