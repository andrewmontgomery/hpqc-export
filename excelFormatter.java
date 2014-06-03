package hpqcExporter;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;
import java.io.PrintStream;
import java.util.ArrayList;

import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class excelFormatter extends WindowAdapter implements WindowListener,
		ActionListener, Runnable {
	private JFrame frame;
	private JTextArea textArea;
	private Thread reader;
	private Thread reader2;
	private boolean quit;

	private final PipedInputStream pin = new PipedInputStream();
	private final PipedInputStream pin2 = new PipedInputStream();

	Thread errorThrower;

	public excelFormatter() {

		frame = new JFrame("HPQC Export-Helper - amontgomery");
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		Dimension frameSize = new Dimension((int) (screenSize.width / 2),
				(int) (screenSize.height / 2));
		int x = (int) (frameSize.width / 2);
		int y = (int) (frameSize.height / 2);
		frame.setBounds(x, y, frameSize.width, frameSize.height);

		textArea = new JTextArea();
		textArea.setEditable(false);

		frame.getContentPane().setLayout(new BorderLayout());
		frame.getContentPane().add(new JScrollPane(textArea),
				BorderLayout.CENTER);

		frame.setVisible(true);

		frame.addWindowListener(this);

		try {
			PipedOutputStream pout = new PipedOutputStream(this.pin);
			System.setOut(new PrintStream(pout, true));
		} catch (java.io.IOException io) {
			textArea.append("Couldn't redirect STDOUT to this console\n"
					+ io.getMessage());
		} catch (SecurityException se) {
			textArea.append("Couldn't redirect STDOUT to this console\n"
					+ se.getMessage());
		}

		try {
			PipedOutputStream pout2 = new PipedOutputStream(this.pin2);
			System.setErr(new PrintStream(pout2, true));
		} catch (java.io.IOException io) {
			textArea.append("Couldn't redirect STDERR to this console\n"
					+ io.getMessage());
		} catch (SecurityException se) {
			textArea.append("Couldn't redirect STDERR to this console\n"
					+ se.getMessage());
		}

		quit = false;

		reader = new Thread(this);
		reader.setDaemon(true);
		reader.start();

		reader2 = new Thread(this);
		reader2.setDaemon(true);
		reader2.start();

	}

	public synchronized void windowClosed(WindowEvent evt) {
		quit = true;
		this.notifyAll(); // Stop ALL threads
		try {
			reader.join(1000);
			pin.close();
		} catch (Exception e) {
		}
		try {
			reader2.join(1000);
			pin2.close();
		} catch (Exception e) {
		}
		System.exit(0);
	}

	public synchronized void windowClosing(WindowEvent evt) {
		frame.setVisible(false);
		frame.dispose();
	}

	public synchronized void actionPerformed(ActionEvent evt) {
		textArea.setText("");
	}

	public synchronized void run() {
		try {
			while (Thread.currentThread() == reader) {
				try {
					this.wait(100);
				} catch (InterruptedException ie) {
				}
				if (pin.available() != 0) {
					String input = this.readLine(pin);
					textArea.append(input);
				}
				if (quit)
					return;
			}

			while (Thread.currentThread() == reader2) {
				try {
					this.wait(100);
				} catch (InterruptedException ie) {
				}
				if (pin2.available() != 0) {
					String input = this.readLine(pin2);
					textArea.append(input);
				}
				if (quit)
					return;
			}
		} catch (Exception e) {
			textArea.append("\nConsole reports an Internal error.");
			textArea.append("The error is: " + e);
		}

		// JUST FOR TESTING (Throw a null pointer after 1 second)
		if (Thread.currentThread() == errorThrower) {
			try {
				this.wait(1000);
			} catch (InterruptedException ie) {
			}
			throw new NullPointerException(
					"Application test: throwing an NullPointerException It should arrive at the console");
		}

	}

	public synchronized String readLine(PipedInputStream in) throws IOException {
		String input = "";
		do {
			int available = in.available();
			if (available == 0)
				break;
			byte b[] = new byte[available];
			in.read(b);
			input = input + new String(b, 0, b.length);
		} while (!input.endsWith("\n") && !input.endsWith("\r\n") && !quit);
		return input;
	}

	public static void main(String[] arg) {
		try {

			new excelFormatter();

			String currentDir = System.getProperty("C:\\export");

			Dimension minsize = new Dimension(500, 500);
			JFrame frame = new JFrame("HPQC Export-Helper");
			frame.setMinimumSize(minsize);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			int preconditionFileRows = 0;

			File folder = new File(currentDir);
			File[] listOfFiles = folder.listFiles();
			ArrayList<File> listOfExcelFiles = new ArrayList<File>();
			File dir = new File("HPQCready");
			dir.mkdir();

			for (int i = 0; i < listOfFiles.length; i++) {

				String extension = "";

				int j = listOfFiles[i].toString().lastIndexOf('.');
				if (j > 0) {
					extension = listOfFiles[i].toString().substring(j + 1);
				}

				if ((listOfFiles[i].isFile()) && extension.equals("xls")) {
					listOfExcelFiles.add(listOfFiles[i]);
					System.out.println(listOfFiles[i]);
				}
			}

			File totalfile = new File(currentDir + "/HPQCReady/HPQC-Total.xls");
			HSSFWorkbook totalworkbook = new HSSFWorkbook();
			totalworkbook.createSheet("Preconditions");
			totalworkbook.createSheet("Execution");
			totalworkbook.createSheet("Validation");
			HSSFRow totalrow = null;

			int size = listOfExcelFiles.size();
			for (int j = 0; j < size; j++) {

				int rt = 0;
				File file = new File(currentDir + "/HPQCReady/"
						+ listOfExcelFiles.get(j).getName());

				String fileNameWithExt = file.getName();
				String fileName = FilenameUtils
						.removeExtension(fileNameWithExt);
				String macro = null;
				String[] rowheader = { "Subject", "Test Name", "Description",
						"Step Name", "Step Description", "Expected Results",
						"Step Type", "Designer", "Project Test Type",
						"Review Status", "Attachments", "Comments",
						"TDT Objectives", "TDT Preconditions", "TDT Execution",
						"TDT Expected Results", "Test Requirement Branch (1)",
						"Test Requirement Branch (2)",
						"Test Requirement Branch (3)",
						"Test Requirement Branch (4)",
						"Test Requirement Branch (5)" };

				String[] sheetNames = { "Preconditions", "Execution",
						"Validation" };

				FileInputStream currentFile = new FileInputStream(
						listOfExcelFiles.get(j));

				HSSFWorkbook current = new HSSFWorkbook(currentFile);
				HSSFWorkbook workbook = new HSSFWorkbook();

				boolean stop = false;
				boolean nonBlankRowFound;
				int c;
				HSSFRow lastRow = null;
				HSSFCell cell4 = null;
				HSSFSheet sheet1 = current.getSheetAt(0);

				while (stop == false) {
					nonBlankRowFound = false;
					lastRow = sheet1.getRow(sheet1.getLastRowNum());
					for (c = lastRow.getFirstCellNum(); c <= lastRow
							.getLastCellNum(); c++) {
						cell4 = lastRow.getCell(c);
						if (cell4 != null
								&& lastRow.getCell(c).getCellType() != HSSFCell.CELL_TYPE_BLANK) {
							nonBlankRowFound = true;
						}
					}
					if (nonBlankRowFound == true) {
						stop = true;
					} else {
						sheet1.removeRow(lastRow);
					}
				}

				while (rt < 3) {

					HSSFSheet sheet = workbook.createSheet(sheetNames[rt]);
					HSSFSheet currentSheet = current.getSheetAt(0);
					HSSFSheet totalsheet = totalworkbook.getSheetAt(0);

					Cell cell = null;
					HSSFRow rowhead = sheet.createRow(0);
					HSSFRow row = null;

					for (int x = 0; 21 > x; x++)

					{

						rowhead.createCell(x).setCellValue(rowheader[x]);

					}

					for (int y = 1; currentSheet.getPhysicalNumberOfRows() > y; y++)

					{

						row = sheet.createRow(y);

						totalrow = totalsheet.createRow(preconditionFileRows);
						preconditionFileRows++;

						int i = 0;

						int found = file.toString().indexOf(".", i);
						int start = file.toString().lastIndexOf("\\") + 1;
						int end = found;
						String project = file.toString().substring(start, end);
						start = end + 1;
						end = file.toString().indexOf(".", start);
						String int_ext = file.toString().substring(start, end);
						start = end + 1;
						end = file.toString().indexOf(".", start);
						String area = file.toString().substring(start, end);
						start = end + 1;
						end = file.toString().indexOf(".", start);
						String testcase = file.toString().substring(start, end);
						start = end + 1;
						end = file.toString().indexOf(".", start);
						String designer = file.toString().substring(start, end);

						if (int_ext.equals("INT"))
							int_ext = "Internal";
						else
							int_ext = "External";

						row.createCell(0).setCellValue(
								project + "\\" + area + "\\" + int_ext + "\\"
										+ testcase);
						totalrow.createCell(0).setCellValue(
								project + "\\" + area + "\\" + int_ext + "\\"
										+ testcase);
						if (y < 10) {
							row.createCell(1)
									.setCellValue(testcase + "_00" + y);
							totalrow.createCell(1).setCellValue(
									testcase + "_00" + y);
						}
						if (y >= 10 && y < 100) {
							row.createCell(1).setCellValue(testcase + "_0" + y);
							totalrow.createCell(1).setCellValue(
									testcase + "_0" + y);
						}
						if (y >= 100 && y < 1000) {
							row.createCell(1).setCellValue(testcase + "_" + y);
							totalrow.createCell(1).setCellValue(
									testcase + "_" + y);
						}

						row.createCell(3).setCellValue("Step " + (rt + 1));
						totalrow.createCell(3).setCellValue("Step " + (rt + 1));

						if (rt == 0) {
							if (currentSheet.getRow(y).getCell(8) != null) {
								cell = currentSheet.getRow(y).getCell(8);
								row.createCell(4).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(4).setCellValue(
										cell.getStringCellValue());
								row.createCell(5).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(5).setCellValue(
										cell.getStringCellValue());
							}
							row.createCell(6).setCellValue("Pre-condition");
							totalrow.createCell(6)
									.setCellValue("Pre-condition");

						}

						else if (rt == 1) {
							if (currentSheet.getRow(y).getCell(9) != null) {
								cell = currentSheet.getRow(y).getCell(9);
								row.createCell(4).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(4).setCellValue(
										cell.getStringCellValue());
								row.createCell(5).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(5).setCellValue(
										cell.getStringCellValue());
							}
							row.createCell(6).setCellValue("Execution");
							totalrow.createCell(6).setCellValue("Execution");
						}

						else if (rt == 2) {
							if (currentSheet.getRow(y).getCell(7) != null) {
								cell = currentSheet.getRow(y).getCell(7);
								row.createCell(4).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(4).setCellValue(
										cell.getStringCellValue());
							}
							if (currentSheet.getRow(y).getCell(10) != null) {
								cell = currentSheet.getRow(y).getCell(10);
								row.createCell(5).setCellValue(
										cell.getStringCellValue());
								totalrow.createCell(5).setCellValue(
										cell.getStringCellValue());
							}
							row.createCell(6).setCellValue("Validation");
							totalrow.createCell(6).setCellValue("Validation");

						}

						row.createCell(7).setCellValue(designer);
						totalrow.createCell(7).setCellValue(designer);
						row.createCell(8).setCellValue("Functional");
						totalrow.createCell(8).setCellValue("Functional");
						row.createCell(9).setCellValue("Draft");
						totalrow.createCell(9).setCellValue("Draft");
						cell = currentSheet.getRow(y).getCell(7);
						row.createCell(12).setCellValue(
								cell.getStringCellValue());
						totalrow.createCell(12).setCellValue(
								cell.getStringCellValue());
						if (currentSheet.getRow(y).getCell(8) != null) {
							cell = currentSheet.getRow(y).getCell(8);
							row.createCell(13).setCellValue(
									cell.getStringCellValue());
							totalrow.createCell(13).setCellValue(
									cell.getStringCellValue());
						}
						if (currentSheet.getRow(y).getCell(9) != null) {
							cell = currentSheet.getRow(y).getCell(9);
							row.createCell(14).setCellValue(
									cell.getStringCellValue());
							totalrow.createCell(14).setCellValue(
									cell.getStringCellValue());
						}
						if (currentSheet.getRow(y).getCell(10) != null) {
							cell = currentSheet.getRow(y).getCell(10);
							row.createCell(15).setCellValue(
									cell.getStringCellValue());
							totalrow.createCell(15).setCellValue(
									cell.getStringCellValue());
						}
						cell = currentSheet.getRow(y).getCell(1);
						row.createCell(16).setCellValue(
								cell.getStringCellValue());
						totalrow.createCell(16).setCellValue(
								cell.getStringCellValue());
						cell = currentSheet.getRow(y).getCell(2);
						row.createCell(17).setCellValue(
								cell.getStringCellValue());
						totalrow.createCell(17).setCellValue(
								cell.getStringCellValue());
						cell = currentSheet.getRow(y).getCell(3);
						row.createCell(18).setCellValue(
								cell.getStringCellValue());
						totalrow.createCell(18).setCellValue(
								cell.getStringCellValue());
						cell = currentSheet.getRow(y).getCell(4);
						row.createCell(19).setCellValue(
								cell.getStringCellValue());
						totalrow.createCell(19).setCellValue(
								cell.getStringCellValue());
						if (currentSheet.getRow(y).getCell(5) != null) {
							cell = currentSheet.getRow(y).getCell(5);

							row.createCell(20).setCellValue(
									cell.getStringCellValue());

							totalrow.createCell(20).setCellValue(
									cell.getStringCellValue());
						}

						else {
							row.createCell(20).setCellValue("");

							totalrow.createCell(20).setCellValue("");
						}

						macro = "Objective:\n"
								+ row.getCell(12).getStringCellValue()
								+ "\n\nPre-condition//Prerequisites\n"
								+ row.getCell(13).getStringCellValue()
								+ "\n\nExecution\n"
								+ row.getCell(14).getStringCellValue()
								+ "\n\nExpected Results\n"
								+ row.getCell(15).getStringCellValue()
								+ "\n\nTest Requirement Branch 1:\n"
								+ row.getCell(16).getStringCellValue()
								+ "\n\nTest Requirement Branch 2:\n"
								+ row.getCell(17).getStringCellValue()
								+ "\n\nTest Requirement Branch 3:\n"
								+ row.getCell(18).getStringCellValue()
								+ "\n\nTest Requirement Branch 4:\n"
								+ row.getCell(19).getStringCellValue()
								+ "\n\nTest Requirement Branch 5:\n"
								+ row.getCell(20).getStringCellValue() + "\n";

						row.createCell(2).setCellValue(macro);
						totalrow.createCell(2).setCellValue(macro);

					}

					for (int z = 0; z < 21; z++)
						sheet.autoSizeColumn(z);

					rt++;
				}

				System.out.println("Completed: " + fileName);
				FileOutputStream fileOut = new FileOutputStream(file);

				workbook.write(fileOut);

				fileOut.close();

			}

			FileOutputStream fileOut2 = new FileOutputStream(totalfile);
			totalworkbook.write(fileOut2);
			fileOut2.close();
			System.out.println("All Done! Ready to export to HPQC!");

		} catch (Exception ex) {

			System.out.println(ex);

		}
	}

}
