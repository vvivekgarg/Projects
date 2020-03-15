
package result;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class ResultCreator {
	Display display;
	Shell shell;
	Text classText;
	Text session;
	Text schoolHeader;
	Text logs;

	public ResultCreator() {
		this.display = new Display();
		(this.shell = new Shell(this.display)).setText("Result Creator");
		final GridLayout gridLayout = new GridLayout(4, true);
		gridLayout.verticalSpacing = 8;
		this.shell.setLayout(gridLayout);
		Label label = new Label(this.shell, 0);
		label.setText("Import File");
		final Text inputFile = new Text(this.shell, 2056);
		GridData gridData = new GridData(256);
		gridData.horizontalSpan = 2;
		inputFile.setLayoutData(gridData);
		final Button importButton = new Button(this.shell, 0);
		final GridData importGridData = new GridData(256);
		importGridData.horizontalSpan = 1;
		importButton.setLayoutData(importGridData);
		importButton.setText("Browse");
		importButton.addSelectionListener(new SelectionListener() {

			@Override
			public void widgetSelected(SelectionEvent arg0) {
				final FileDialog fd = new FileDialog(shell, SWT.OPEN);
				fd.setText("Open");
				fd.setFilterPath("C:/");
				final String[] filterExt = { "*.xls" };
				fd.setFilterExtensions(filterExt);
				final String selected = fd.open();
				if (selected != null && selected.endsWith(".xls")) {
					inputFile.setText(selected);
				}
			}

			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
				// TODO Auto-generated method stub

			}
		});
		label = new Label(this.shell, 0);
		label.setText("School Header: ");
		this.schoolHeader = new Text(this.shell, 2052);
		gridData = new GridData(256);
		gridData.horizontalSpan = 3;
		this.schoolHeader.setLayoutData(gridData);
		this.schoolHeader.setText("");
		label = new Label(this.shell, 0);
		label.setText("Result Year Session: ");
		this.session = new Text(this.shell, 2052);
		gridData = new GridData(256);
		gridData.horizontalSpan = 3;
		this.session.setLayoutData(gridData);
		this.session.setText("");
		label = new Label(this.shell, 0);
		label.setText("Class: ");
		this.classText = new Text(this.shell, 2052);
		gridData = new GridData(256);
		gridData.horizontalSpan = 3;
		this.classText.setLayoutData(gridData);
		this.classText.setText("");
		label = new Label(this.shell, 0);
		final Button resultCreate = new Button(this.shell, 0);
		gridData = new GridData(256);
		gridData.horizontalSpan = 2;
		resultCreate.setLayoutData(gridData);
		resultCreate.setText("Generate Result");
		label = new Label(this.shell, 0);
		label = new Label(this.shell, 0);
		label.setText("Logs:");
		this.logs = new Text(this.shell, 2882);
		gridData = new GridData(272);
		gridData.horizontalSpan = 3;
		gridData.grabExcessVerticalSpace = true;
		gridData.grabExcessHorizontalSpace = true;
		this.logs.setLayoutData(gridData);
		this.logs.setText("");
		label = new Label(this.shell, 0);
		label = new Label(this.shell, 0);
		label.setText("Created By : Vivek Garg ©");
		resultCreate.addSelectionListener(new SelectionListener() {

			@Override
			public void widgetSelected(SelectionEvent arg0) {

				final MessageBox messageBox = new MessageBox(shell, 40);
				messageBox.setText("Warning");
				if (classText.getText() == null || classText.getText().isEmpty()) {
					messageBox.setMessage("Please provide class information.");
					messageBox.open();
				} else if (session.getText() == null || session.getText().isEmpty()) {
					messageBox.setMessage("Please provide Session information.");
					messageBox.open();
				} else if (schoolHeader.getText() == null || schoolHeader.getText().isEmpty()) {
					messageBox.setMessage("Please provide School Header information.");
					messageBox.open();
				} else if (inputFile.getText() == null || inputFile.getText().isEmpty()) {
					messageBox.setMessage("Please provide input file.");
					messageBox.open();
				} else {
					try {
						logs.setText("");
						final Reader reader = new Reader();
						final Map<String, String> inputValues = new TreeMap<String, String>();
						inputValues.put("session", session.getText());
						inputValues.put("class", classText.getText());
						inputValues.put("schoolHeader", schoolHeader.getText());
						final FileDialog dialog = new FileDialog(shell, 8192);
						dialog.setFilterExtensions(new String[] { "*.xls" });
						dialog.setFilterPath("c:\\");
						final String saveFile = dialog.open();
						if (saveFile == null || saveFile.trim().isEmpty()) {
							messageBox.setMessage("Please specify corerct file for saving result.");
							messageBox.open();
						} else {
							logs.append("Reading Input file. \n");
							final List<Student> students = reader.readFile(inputFile.getText());
							logs.append("Successfully read information of " + students.size() + " students.\n");
							final FileOutputStream out = new FileOutputStream(saveFile);
							final Workbook workbook = new HSSFWorkbook();
							new ExcelUtils(workbook);
							Sheet sheet = workbook.createSheet("Student's Result");
							logs.append("Started writing student wise details.\n");
							final StudentResultSheet header = new StudentResultSheet(workbook);
							Integer rowNum = 0;
							for (final Student student : students) {
								logs.append("\nWriting Details for : " + student.getName());
								rowNum = header.createStudentResultSheets(sheet, student, inputValues, rowNum);
							}
							logs.append("\nSuccessfully written student wise result.\n");
							sheet = workbook.createSheet("Final Result");
							logs.append("Started writing final results.\n");
							final FinalResultSheet finalResult = new FinalResultSheet(
									reader.getSheetNames(inputFile.getText()));
							finalResult.createFinalResultSheets(sheet, students, inputValues);
							logs.append("Successfully written final result.\n");
							workbook.write(out);
							out.close();
							messageBox.setMessage("File Generated Successfully.");
							messageBox.open();
						}
					} catch (IOException e) {
						messageBox.setMessage("File is already opened or Incorrect file has been provided.");
						messageBox.open();
						e.printStackTrace();
					} catch (Exception e2) {
						e2.printStackTrace();
						messageBox.setMessage(e2.getMessage());
						messageBox.open();
					}
				}
			}

			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
			}
		});
		this.shell.setSize(800, 500);
		this.shell.open();
		while (!this.shell.isDisposed()) {
			if (!this.display.readAndDispatch()) {
				this.display.sleep();
			}
		}
		this.display.dispose();
	}

	public static void main(final String[] args) {
		new ResultCreator();
	}
}
