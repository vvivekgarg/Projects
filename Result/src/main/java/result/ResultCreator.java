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
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class ResultCreator {

	Display display;
	Shell shell;

	public ResultCreator() {
		display = new Display();
		shell = new Shell(display);
		shell.setText("Result Creator");

		GridLayout gridLayout = new GridLayout(4, true);
		gridLayout.verticalSpacing = 8;

		shell.setLayout(gridLayout);

		// Title
		Label label = new Label(shell, SWT.NULL);

		label.setText("Import File");
		Text inputFile = new Text(shell, SWT.READ_ONLY | SWT.BORDER);
		GridData gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		gridData.horizontalSpan = 2;
		inputFile.setLayoutData(gridData);

		Button importButton = new Button(shell, SWT.NONE);
		GridData importGridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		importGridData.horizontalSpan = 1;
		importButton.setLayoutData(importGridData);
		importButton.setText("Browse");
		importButton.addSelectionListener(new SelectionListener() {

			@Override
			public void widgetSelected(SelectionEvent arg0) {
				FileDialog fd = new FileDialog(shell, SWT.OPEN);
				fd.setText("Open");
				fd.setFilterPath("C:/");
				String[] filterExt = { "*.xls" };
				fd.setFilterExtensions(filterExt);
				String selected = fd.open();
				if (selected != null && selected.endsWith(".xls"))
					inputFile.setText(selected);

			}

			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
				// TODO Auto-generated method stub

			}
		});

		label = new Label(shell, SWT.NULL);
		label.setText("School Header: ");

		Text schoolHeader = new Text(shell, SWT.SINGLE | SWT.BORDER);
		gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		gridData.horizontalSpan = 3;
		schoolHeader.setLayoutData(gridData);
		schoolHeader.setText("");

		// Author(s)
		label = new Label(shell, SWT.NULL);
		label.setText("Result Year Session: ");

		Text session = new Text(shell, SWT.SINGLE | SWT.BORDER);
		gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		gridData.horizontalSpan = 3;
		session.setLayoutData(gridData);
		session.setText("");

		label = new Label(shell, SWT.NULL);
		label.setText("Class: ");

		Text classText = new Text(shell, SWT.SINGLE | SWT.BORDER);
		gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		gridData.horizontalSpan = 3;
		classText.setLayoutData(gridData);
		classText.setText("");

		label = new Label(shell, SWT.NULL);
		Button resultCreate = new Button(shell, SWT.NONE);
		gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL);
		gridData.horizontalSpan = 2;
		resultCreate.setLayoutData(gridData);
		resultCreate.setText("Generate Result");
		label = new Label(shell, SWT.NULL);
		label = new Label(shell, SWT.NULL);
		label.setText("Logs:");

		Text logs = new Text(shell, SWT.WRAP | SWT.MULTI | SWT.BORDER | SWT.H_SCROLL | SWT.V_SCROLL);
		gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL | GridData.VERTICAL_ALIGN_FILL);
		gridData.horizontalSpan = 3;
		gridData.grabExcessVerticalSpace = true;
		gridData.grabExcessHorizontalSpace = true;

		logs.setLayoutData(gridData);
		logs.setText("");
		resultCreate.addSelectionListener(new SelectionListener() {

			@Override
			public void widgetSelected(SelectionEvent arg0) {

				MessageBox messageBox = new MessageBox(shell, SWT.ICON_WARNING | SWT.OK);
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
				}

				else if (inputFile.getText() == null || inputFile.getText().isEmpty()) {
					messageBox.setMessage("Please provide input file.");
					messageBox.open();
				}

				else {

					try {
						logs.setText("");
						Reader reader = new Reader();
						Map<String, String> inputValues = new TreeMap<>();
						inputValues.put("session", session.getText());
						inputValues.put("class", classText.getText());
						inputValues.put("schoolHeader", schoolHeader.getText());

						FileDialog dialog = new FileDialog(shell, SWT.SAVE);
						dialog.setFilterExtensions(new String[] { "*.xls" }); // Windows
						// wild
						// cards
						dialog.setFilterPath("c:\\"); // Windows path
						String saveFile = dialog.open();
						if (saveFile == null || saveFile.trim().isEmpty()) {
							messageBox.setMessage("Please specify corerct file for saving result.");
							messageBox.open();
						} else {
							logs.append("Reading Input file. \n");
							List<Student> students = reader.readFile(inputFile.getText());
							logs.append("Successfully read information of " + students.size() + " students.\n");

							FileOutputStream out = new FileOutputStream(saveFile);
							Workbook workbook = new HSSFWorkbook();
							// Ininitialize Styles
							new ExcelUtils(workbook);

							// Student wise result
							Sheet sheet = workbook.createSheet("Student's Result");
							logs.append("Started writing student wise details.\n");
							StudentResultSheet header = new StudentResultSheet(workbook);
							header.createStudentResultSheets(sheet, students, inputValues);
							logs.append("Successfully written student wise result.\n");
							// FInal Result
							sheet = workbook.createSheet("Final Result");
							logs.append("Started writing final results.\n");
							FinalResultSheet finalResult = new FinalResultSheet(
									reader.getSheetNames(inputFile.getText()));
							finalResult.createFinalResultSheets(sheet, students, inputValues);
							logs.append("Successfully written final result.\n");

							workbook.write(out);
							out.close();
							messageBox.setMessage("File Generated Successfully.");
							messageBox.open();
						}
					} catch (Exception e) {
						messageBox.setMessage("File is already opened or Incorrect file has been provided.");
						messageBox.open();
						e.printStackTrace();
					}
				}
			}

			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
			}
		});

		shell.setSize(800, 500);
		// shell.pack();
		shell.open();
		// Set up the event loop.
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				// If no more entries in event queue
				display.sleep();
			}
		}

		display.dispose();
	}

	public static void main(String[] args) {
		new ResultCreator();
	}
}
