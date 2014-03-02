package Automation;

//BEGIN custom imports
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;

import com.sun.jna.Native;
import com.sun.jna.PointerType;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.RECT;
import com.sun.jna.win32.W32APIOptions;
import jxl.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Orientation;
import jxl.format.UnderlineStyle;
import java.awt.AWTException;
import java.awt.Container;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.InputEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.HeadlessException;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.Robot;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;

/**
 * <p>
 * <b>Reporting Tool Program: </b>Populates results in the MS-Excel Spreadsheet Report (Template)
 * from the WPSS tool, CSE Pro tool and W3C tool
 * 
 * </p>
 * <b>Description: </b> Permission is hereby granted, free of charge, to anyone
 * obtaining a copy of this software and associated documentation files (the
 * "Software"). Therefore, the author reserve limitations and rights to modify,
 * merge, publish, sublicense and sell. Copyright has been reserved to Matrixx
 * Hi-Tech Inc. The Software can be distribute under the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 * 
 * @author Gaston Langlois - Environment Canada - Matrixx Hi-Tech Inc.
 * @since 2012/02/15
 */

public class Script5 {
	public interface User32 extends W32APIOptions {
		public static final String SHELL_TRAY_WND = "Shell_TrayWnd";
		public static final int WM_COMMAND = 0x111;
		public static final int MIN_ALL = 0x1a3;
		public static final int MIN_ALL_UNDO = 0x1a0;
		User32 INSTANCE = (User32) Native.loadLibrary("user32", User32.class);

		HWND GetForegroundWindow();

		int GetWindowTextA(PointerType hWnd, byte[] lpString, int nMaxCount);

		User32 instance = (User32) Native.loadLibrary("user32", User32.class,
				DEFAULT_OPTIONS);

		HWND FindWindow(String winClass, String title);

		long SendMessageA(HWND hWnd, int msg, int num1, int num2);

		boolean IsWindowVisible(int hWnd);

		int GetWindowRect(int hWnd, RECT r);

		void GetWindowTextA(int hWnd, byte[] buffer, int buflen);

		int GetTopWindow(int hWnd);

		int GetWindow(int hWnd, int flag);

		final int GW_HWNDNEXT = 2;
	}

	@SuppressWarnings({ "resource", "deprecation" })
	public static void main(String[] args) throws InterruptedException,
			RowsExceededException, WriteException, IOException, JXLException,
			AWTException, HeadlessException, UnsupportedFlavorException {

		Robot robot = new Robot();
		File file = new File("C://QA-WS-Tool//filename.txt");
		file.delete();
		// It will make sure nvda is closed
		Runtime.getRuntime()
				.exec("wmic process where name=\"nvda.exe\" delete");

		String[] StringMove = { "cmd.exe", "/c", "ECHO %USERNAME%" };
		Process ProcessMove = Runtime.getRuntime().exec(StringMove);
		BufferedReader VarMove = new BufferedReader(new InputStreamReader(
				ProcessMove.getInputStream()));
		String temp = "";
		String Username = "";
		while ((temp = VarMove.readLine()) != null) {
			Thread.sleep(1);
			Username = temp;
		}
		VarMove.close();

		// CHECK WHETHER INPUT FILE EXIST (not for .txt files)
		String fileName = "C:///QA-WS-Tool//Input.xls";
		File file1 = new File(fileName);

		// try to rename the file with the same name
		File sameFileName = new File(fileName);
		if (file1.renameTo(sameFileName)) {
			// Validate URL in the input file
			validateURLs.main(args);
			// System.out.println("file is closed");
		} else {
			// if the file didnt accept the renaming operation
			JOptionPane pane1 = new JOptionPane(
					"The following file is not available, already in use or missing: "
							+ "C:\\QA-WS-Tool\\Input.xls\n\n"
							+ "Please make sure to Close and Save your INPUT file before starting then try again\n\n"
							+ "Press Ok to continue");
			JDialog d1 = pane1.createDialog((JFrame) null,
					"C:\\QA-WS-Tool\\Input.xls file is unavailable or missing");
			d1.setLocation(400, 500);
			d1.setVisible(true);
			System.out
					.println("Input file C:\\QA-WS-Tool\\Input.xls unavailable or missing");
			System.exit(0);
		}

		// Close Internet Explore if opened
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
		robot.delay(2000);

		// get the taskbar's window handle
		HWND shellTrayHwnd = User32.instance.FindWindow(User32.SHELL_TRAY_WND,
				null);
		// Restore windows
		// use it to minimize all windows
		User32.instance.SendMessageA(shellTrayHwnd, User32.WM_COMMAND,
				User32.MIN_ALL, 0);

		final JFrame f1 = new JFrame("QA Web Standards Reporting Tool V1.1");
		// f.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		//f1.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		f1.setSize(400, 200);
		f1.setLocation(300, 300);

		f1.addWindowListener(new WindowAdapter() {
		public void windowClosing(WindowEvent we) {
		
		System.out.println("Please note that all Java session open are now closed");	
		try {
			Runtime.getRuntime().exec("taskkill /F /IM javaw.exe ");
		} catch (IOException e) {
			System.out.println("taskkill /F /IM javaw.exe was executed\n"+e);
			e.printStackTrace();
		}
		System.exit(0);
		//new ClosingFrame();
		//f1.setVisible(false);
		}
		});
		// String TitleContent = "";
		JPanel entreePanel = new JPanel();
		final ButtonGroup entreeGroup = new ButtonGroup();
		// Add a text message to select the tool to run
		String text = "Please select the tool to run:";
		text += "\n";

		JLabel testLabel = new JLabel(text);
		entreePanel.add(testLabel);
		JRadioButton radioButton;
		entreePanel.add(radioButton = new JRadioButton("WPSS"));
		radioButton.setActionCommand("WPSS");
		radioButton.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup.add(radioButton);
		entreePanel.add(radioButton = new JRadioButton("CSE"));
		radioButton.setActionCommand("CSE");
		radioButton.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup.add(radioButton);
		entreePanel.add(radioButton = new JRadioButton("W3C", true));
		radioButton.setActionCommand("W3C");
		radioButton.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup.add(radioButton);
		testLabel.setFont(new java.awt.Font("Arial", 0, 14));
		testLabel.add(Box.createHorizontalStrut(5));
		testLabel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

		// final JPanel condimentsPanel = new JPanel();
		// condimentsPanel.add(new JCheckBox("Direct Input"));
		// condimentsPanel.add(new JCheckBox("Live"));
		JPanel entreePanel2 = new JPanel();
		final ButtonGroup entreeGroup2 = new ButtonGroup();
		// Add a text message to select the tool to run
		String text2 = "Select the type of Source code:      ";
		text2 += "\n";
		JLabel testLabel2 = new JLabel(text2);

		// customize radio button input
		entreePanel2.add(testLabel2);
		JRadioButton radioButton2;
		entreePanel2
				.add(radioButton2 = new JRadioButton("Direct Input", false));
		radioButton2.setActionCommand("Direct Input");
		radioButton2.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup2.add(radioButton2);
		entreePanel2.add(radioButton2 = new JRadioButton("Url", true));
		radioButton2.setActionCommand("Url");
		// // Preselect the Live radio button
		// radioButton2.setSelected(true);
		radioButton2.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup2.add(radioButton2);
		testLabel2.add(Box.createHorizontalStrut(5));
		testLabel2.setFont(new java.awt.Font("Arial", 0, 14));

		JPanel orderPanel = new JPanel();
		JButton orderButton = new JButton("Submit");
		orderPanel.add(orderButton);
		orderPanel.setFont(new java.awt.Font("Arial", 0, 14));

		Container content = f1.getContentPane();
		content.setLayout(new GridLayout(3, 1));
		content.add(entreePanel);
		// content.add(condimentsPanel);
		content.add(entreePanel2);
		content.add(orderPanel);

		orderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				String entree = entreeGroup.getSelection().getActionCommand();
				PrintWriter writer = null;
				try {
					writer = new PrintWriter("C://QA-WS-Tool//filename.txt",
							"UTF-8");
				} catch (FileNotFoundException e) {
					// All cells modified/added. Now write out the workbook
					e.printStackTrace();
					System.out.println("File not found in C://QA-WS-Tool//filename.txt\n"+e);
				} catch (UnsupportedEncodingException e) {
					// All cells modified/added. Now write out the workbook
					e.printStackTrace();
					System.out.println("UnsupportedEncodingException/n"+e);
				}
				writer.println(entree);
				String entree2 = entreeGroup2.getSelection().getActionCommand();
				writer.println(entree2);
				// System.out.println(entree2);
				// Component[] components = condimentsPanel.getComponents();
				// for (int i = 0; i < components.length; i++) {
				// JCheckBox cb = (JCheckBox) components[i];
				// if (cb.isSelected()) {
				// writer.println(cb.getText());
				// //System.out.println("With " + cb.getText());
				// }
				// }
				writer.close();
				f1.setVisible(false);
			}
		});
		f1.setVisible(true);

		File f = new File("C://QA-WS-Tool//filename.txt");
		while (f.exists() == false) {
			robot.delay(3000);
		}

		String Tool = "";
		String SourceCode = "";
		try {
			// Read and copy file
			FileInputStream inputFile = new FileInputStream(
					"C:\\QA-WS-Tool\\filename.txt");
			InputStreamReader frSource = new InputStreamReader(inputFile,
					"UTF-8");
			// FileReader frSource = new
			// FileReader("C:\\QA-WS-Tool\\filename.txt");
			BufferedReader brSource = new BufferedReader(frSource);
			String ReadCurrentLineSource;
			int linenumberSource = 0;
			// Reading WPSS results
			while ((ReadCurrentLineSource = brSource.readLine()) != null) {
				linenumberSource = linenumberSource + 1;
				if (linenumberSource == 1) {
					// System.out.println("Line1="+ReadCurrentLineSource);
					Tool = ReadCurrentLineSource;
				}
				if (linenumberSource == 2) {
					// System.out.println("Line2="+ReadCurrentLineSource);
					SourceCode = ReadCurrentLineSource;
				}
			}
			// System.out.println(contentSource);
		} // End of Read and Copy file

		catch (Exception ex) {
			// All cells modified/added. Now write out the workbook
		}

		// ////////////////////////////////////////////////////////
		// KeyListenerTester.main(null);
		// ////////////////////////////////////////////////////////

		// CHECK WHETHER INPUT FILE EXIST (not for .txt files)
		String fileName1 = "C:///QA-WS-Tool//Output-" + Tool + ".xls";
		File file11 = new File(fileName1);

		// try to rename the file with the same name
		File sameFileName1 = new File(fileName1);
		if (file11.renameTo(sameFileName1)) {
			// if the file is renamed
			// System.out.println("file is closed");
		} else {
			// if the file didnt accept the renaming operation
			JOptionPane pane1 = new JOptionPane(
					"The following file is not available, already in use or missing: "
							+ "C:\\QA-WS-Tool\\Output-"
							+ Tool
							+ ".xls\n\n"
							+ "Please make sure to prepare and close your INPUT file before starting then try again\n\n"
							+ "Press Ok to continue");
			JDialog d1 = pane1.createDialog((JFrame) null,
					"C:\\QA-WS-Tool\\Output-" + Tool
							+ ".xls file is unavailable or missing");
			d1.setLocation(400, 500);
			d1.setVisible(true);
			System.out.println("Input file C:\\QA-WS-Tool\\Output-" + Tool
					+ ".xls unavailable or missing");
			System.exit(0);
		}

		System.out.println("Input file copied from C:\\QA-WS-Tool\\Input.xls");
		Workbook workbook = Workbook.getWorkbook(new File(
				"C://QA-WS-Tool//Input.xls"));
		WritableWorkbook copy = Workbook.createWorkbook(new File(
				"C://QA-WS-Tool//Output-" + Tool + ".xls"), workbook);
		// Sheet2 is the Accessibility Tab in the spreadsheet
		WritableSheet sheet2 = copy.getSheet(2);

		// Sheet4 is the Accessibility Tab in the spreadsheet
		WritableSheet sheet4 = copy.getSheet(4);
		// Set arial9formatNoBold
		WritableFont arial9fontNoBold = new WritableFont(WritableFont.ARIAL, 9,
				WritableFont.NO_BOLD);
		WritableCellFormat arial9formatNoBold = new WritableCellFormat(
				arial9fontNoBold);
		arial9formatNoBold.setAlignment(Alignment.CENTRE);
		arial9formatNoBold.setOrientation(Orientation.HORIZONTAL);
		arial9formatNoBold.setShrinkToFit(false);
		arial9formatNoBold.setWrap(true);
		arial9formatNoBold.setBorder(Border.ALL, BorderLineStyle.THIN);

		// arial9formatNoBold
		WritableFont arial9fontBold = new WritableFont(WritableFont.ARIAL, 9,
				WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE,
				Colour.BLACK);
		WritableCellFormat arial9formatBold = new WritableCellFormat(
				arial9fontBold);
		// arial9formatBold.setBackground(Colour.GREY_25_PERCENT);
		arial9formatBold.setAlignment(Alignment.CENTRE);
		arial9formatBold.setOrientation(Orientation.HORIZONTAL);
		arial9formatBold.setShrinkToFit(false);
		arial9formatBold.setWrap(true);
		arial9formatBold.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Close all session in excel
		// Runtime.getRuntime().exec("taskkill /F /IM excel.exe ");
		System.out.println("Output file created in "
				+ "C:\\QA-WS-Tool\\Output-" + Tool + ".xls");

		// Add a list data validations
		// ArrayList<String> al = new ArrayList<String>();
		// al.add("Pass");
		// al.add("Fail");
		// al.add("N/A");

		// Initialization of variables
		// Declare array for WCAG2 Sufficient Techniques row in the spreasheet
		int[] WCAG2RowArray = new int[38];

		WCAG2RowArray[0] = 7;
		WCAG2RowArray[1] = 9;
		WCAG2RowArray[2] = 10;
		WCAG2RowArray[3] = 11;
		WCAG2RowArray[4] = 12;
		WCAG2RowArray[5] = 13;
		WCAG2RowArray[6] = 15;
		WCAG2RowArray[7] = 16;
		WCAG2RowArray[8] = 17;
		WCAG2RowArray[9] = 19;
		WCAG2RowArray[10] = 20;
		WCAG2RowArray[11] = 21;
		WCAG2RowArray[12] = 22;
		WCAG2RowArray[13] = 23;
		WCAG2RowArray[14] = 25;
		WCAG2RowArray[15] = 26;
		WCAG2RowArray[16] = 28;
		WCAG2RowArray[17] = 29;
		WCAG2RowArray[18] = 31;
		WCAG2RowArray[19] = 33;
		WCAG2RowArray[20] = 34;
		WCAG2RowArray[21] = 35;
		WCAG2RowArray[22] = 36;
		WCAG2RowArray[23] = 37;
		WCAG2RowArray[24] = 38;
		WCAG2RowArray[25] = 39;
		WCAG2RowArray[26] = 41;
		WCAG2RowArray[27] = 42;
		WCAG2RowArray[28] = 44;
		WCAG2RowArray[29] = 45;
		WCAG2RowArray[30] = 46;
		WCAG2RowArray[31] = 47;
		WCAG2RowArray[32] = 49;
		WCAG2RowArray[33] = 50;
		WCAG2RowArray[34] = 51;
		WCAG2RowArray[35] = 52;
		WCAG2RowArray[36] = 54;
		WCAG2RowArray[37] = 55;

		// Initialization of variables
		// Declare array for Interoperability checkpoints in the spreasheet
		int[] InteropRowArray = new int[35];

		InteropRowArray[0] = 7;
		InteropRowArray[1] = 8;
		InteropRowArray[2] = 9;
		InteropRowArray[3] = 10;
		InteropRowArray[4] = 11;
		InteropRowArray[5] = 12;
		InteropRowArray[6] = 13;
		InteropRowArray[7] = 14;
		InteropRowArray[8] = 15;
		InteropRowArray[9] = 16;
		InteropRowArray[10] = 17;
		InteropRowArray[11] = 18;
		InteropRowArray[12] = 19;
		InteropRowArray[13] = 20;
		InteropRowArray[14] = 21;
		InteropRowArray[15] = 22;
		InteropRowArray[16] = 23;
		InteropRowArray[17] = 24;
		InteropRowArray[18] = 25;
		InteropRowArray[19] = 26;
		InteropRowArray[20] = 27;
		InteropRowArray[21] = 28;
		InteropRowArray[22] = 29;
		InteropRowArray[23] = 30;
		InteropRowArray[24] = 31;
		InteropRowArray[25] = 32;
		InteropRowArray[26] = 33;
		InteropRowArray[27] = 34;
		InteropRowArray[28] = 36;
		InteropRowArray[29] = 38;
		InteropRowArray[30] = 39;
		InteropRowArray[31] = 41;
		InteropRowArray[32] = 42;
		InteropRowArray[33] = 43;
		InteropRowArray[34] = 44;

		// Declare array for WCAG2 Sufficient Techniques Strings in the
		// Spreadsheet
		String[] WCAG2StringArray = new String[38];
		WCAG2StringArray[0] = "1.1.1";
		WCAG2StringArray[1] = "1.2.1";
		WCAG2StringArray[2] = "1.2.2";
		WCAG2StringArray[3] = "1.2.3";
		WCAG2StringArray[4] = "1.2.4";
		WCAG2StringArray[5] = "1.2.5";
		WCAG2StringArray[6] = "1.3.1";
		WCAG2StringArray[7] = "1.3.2";
		WCAG2StringArray[8] = "1.3.3";
		WCAG2StringArray[9] = "1.4.1";
		WCAG2StringArray[10] = "1.4.2";
		WCAG2StringArray[11] = "1.4.3";
		WCAG2StringArray[12] = "1.4.4";
		WCAG2StringArray[13] = "1.4.5";
		WCAG2StringArray[14] = "2.1.1";
		WCAG2StringArray[15] = "2.1.2";
		WCAG2StringArray[16] = "2.2.1";
		WCAG2StringArray[17] = "2.2.2";
		WCAG2StringArray[18] = "2.3.1";
		WCAG2StringArray[19] = "2.4.1";
		WCAG2StringArray[20] = "2.4.2";
		WCAG2StringArray[21] = "2.4.3";
		WCAG2StringArray[22] = "2.4.4";
		WCAG2StringArray[23] = "2.4.5";
		WCAG2StringArray[24] = "2.4.6";
		WCAG2StringArray[25] = "2.4.7";
		WCAG2StringArray[26] = "3.1.1";
		WCAG2StringArray[27] = "3.1.2";
		WCAG2StringArray[28] = "3.2.1";
		WCAG2StringArray[29] = "3.2.2";
		WCAG2StringArray[30] = "3.2.3";
		WCAG2StringArray[31] = "3.2.4";
		WCAG2StringArray[32] = "3.3.1";
		WCAG2StringArray[33] = "3.3.2";
		WCAG2StringArray[34] = "3.3.3";
		WCAG2StringArray[35] = "3.3.4";
		WCAG2StringArray[36] = "4.1.1";
		WCAG2StringArray[37] = "4.1.2";
		int Url_ToBeTest = 0;
		for (int i = 1; i < 69; i++) {
		//for (int i = 42; i < 69; i++) {
			Url_ToBeTest = i;
			// Convert Url_ToBeTest in String 
			String Url_Number= "";
			String string = Integer.toString(Url_ToBeTest);
			char[] charArr3 = string.toCharArray();
			if (charArr3.length == 2) {
				Url_Number = "0";
			}
			if (charArr3.length == 1) {
				Url_Number = "00";
			}
			for (int i3 = 0; i3 < charArr3.length; i3++) {
				String Data = String.copyValueOf(charArr3, i3, 1);
				if (Data.contentEquals("0")) {
					Url_Number = Url_Number + "0";
				}
				if (Data.contentEquals("1")) {
					Url_Number = Url_Number + "1";
				}
				if (Data.contentEquals("2")) {
					Url_Number = Url_Number + "2";
				}
				if (Data.contentEquals("3")) {
					Url_Number = Url_Number + "3";
				}
				if (Data.contentEquals("4")) {
					Url_Number = Url_Number + "4";
				}
				if (Data.contentEquals("5")) {
					Url_Number = Url_Number + "5";
				}
				if (Data.contentEquals("6")) {
					Url_Number = Url_Number + "6";
				}
				if (Data.contentEquals("7")) {
					Url_Number = Url_Number + "7";
				}
				if (Data.contentEquals("8")) {
					Url_Number = Url_Number + "8";
				}
				if (Data.contentEquals("9")) {
					Url_Number = Url_Number + "9";
				}
			}
			
			
			// Initialize variables for each URL and WCAG
			String WPSS_Version = "";
			String TitleStripped = "";
			String path = "C:\\QA-WS-Tool\\WPSS_Tool\\results\\WorkingStorage_acc.txt";
			String path2 = "C:\\QA-WS-Tool\\WPSS_Tool\\results\\WorkingStorage_link.txt";
			String path3 = "C:\\Users\\"
					+ Username
					+ "\\AppData\\Roaming\\AI Internet Solutions\\CSE HTML Validator\\14.0\\batchreport1.html";
			String path4 = "C:\\QA-WS-Tool\\WPSS_Tool\\results\\WorkingStorage_int.txt";
			String WorkingStorage = "C:\\QA-WS-Tool\\WPSS_Tool\\results\\WorkingStorage";
			WritableCell cell = sheet2.getWritableCell(Url_ToBeTest, 5);
			WritableCell cell2 = sheet2.getWritableCell(Url_ToBeTest, 3);

			boolean AnalysisCompleted = false;
			boolean WCAG2_Error2_4_2 = false;

			// Get the AdditionalInfo content
			String AdditionalInfo = cell2.getContents().trim();

			// Get page URL content
			String str = cell.getContents().trim();
			int len = str.length();
			if (len < 6) {
				break;
			}

			int PositionOfSlash = str.lastIndexOf("/");
			int PositionOfWebInfo = PositionOfSlash + 1;
			char[] charArray1 = str.toCharArray();
			for (int k = PositionOfWebInfo; k < len; k++) {
				// System.out.print(charArray[k]);
			}
			String.valueOf(charArray1);
			String.valueOf(charArray1);
			String.copyValueOf(charArray1, 0, PositionOfWebInfo);
			String.copyValueOf(charArray1, PositionOfWebInfo, len
					- PositionOfWebInfo);

			if (Tool.contains("CSE") && (SourceCode.contains("Direct Input"))) {
				// Copy URL content in the Clipboard
				Toolkit toolkit = Toolkit.getDefaultToolkit();
				Clipboard clipboard = toolkit.getSystemClipboard();
				StringSelection strSel = new StringSelection(cell.getContents());
				clipboard.setContents(strSel, null);
				// Prompt Tester you require to follow manual steps
				// (if a sequence of steps need to be done to get to this page)
				if ((AdditionalInfo.contains("[Additional Information]") == false)
						&& (AdditionalInfo.contains("Skip") == false)) {
					LaunchFirefox.main(args);
					JOptionPane pane = new JOptionPane(
							"\n"
									+ "Following instruction need to be done to get to this page\n\n"
									+ AdditionalInfo
									+ "\n\n"
									+ "Press Ok to continue when manual steps are completed");
					JDialog d = pane.createDialog((JFrame) null,
							"Manual steps is required");
					d.setLocation(400, 500);
					d.setVisible(true);
					// Press View Source from Firefox
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_U);
					robot.delay(300);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_U);

					// Press Ctrl+S to save the html file and press save
					// button
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_S);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_S);
					robot.delay(2000);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_SEMICOLON);
					robot.delay(100);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_SEMICOLON);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_BACK_SLASH);
					robot.keyRelease(KeyEvent.VK_BACK_SLASH);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_E);
					robot.keyRelease(KeyEvent.VK_E);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_P);
					robot.keyRelease(KeyEvent.VK_P);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_BACK_SLASH);
					robot.keyRelease(KeyEvent.VK_BACK_SLASH);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_R);
					robot.keyRelease(KeyEvent.VK_R);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_SUBTRACT);
					robot.keyRelease(KeyEvent.VK_SUBTRACT);
					robot.delay(100);
					// ///////////////
					String s = Integer.toString(Url_ToBeTest);
					char[] charArr2 = s.toCharArray();
					if (charArr2.length == 2) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30); 
					}
					if (charArr2.length == 1) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
					}
					for (int i1 = 0; i1 < charArr2.length; i1++) {
						String Data = String.copyValueOf(charArr2, i1, 1);
						if (Data.contentEquals("0")) {
							robot.keyPress(KeyEvent.VK_0);
							robot.keyRelease(KeyEvent.VK_0);
							robot.delay(30);
						}
						if (Data.contentEquals("1")) {
							robot.keyPress(KeyEvent.VK_1);
							robot.keyRelease(KeyEvent.VK_1);
							robot.delay(30);
						}
						if (Data.contentEquals("2")) {
							robot.keyPress(KeyEvent.VK_2);
							robot.keyRelease(KeyEvent.VK_2);
							robot.delay(30);
						}
						if (Data.contentEquals("3")) {
							robot.keyPress(KeyEvent.VK_3);
							robot.keyRelease(KeyEvent.VK_3);
							robot.delay(30);
						}
						if (Data.contentEquals("4")) {
							robot.keyPress(KeyEvent.VK_4);
							robot.keyRelease(KeyEvent.VK_4);
							robot.delay(30);
						}
						if (Data.contentEquals("5")) {
							robot.keyPress(KeyEvent.VK_5);
							robot.keyRelease(KeyEvent.VK_5);
							robot.delay(30);
						}
						if (Data.contentEquals("6")) {
							robot.keyPress(KeyEvent.VK_6);
							robot.keyRelease(KeyEvent.VK_6);
							robot.delay(30);
						}
						if (Data.contentEquals("7")) {
							robot.keyPress(KeyEvent.VK_7);
							robot.keyRelease(KeyEvent.VK_7);
							robot.delay(30);
						}
						if (Data.contentEquals("8")) {
							robot.keyPress(KeyEvent.VK_8);
							robot.keyRelease(KeyEvent.VK_8);
							robot.delay(30);
						}
						if (Data.contentEquals("9")) {
							robot.keyPress(KeyEvent.VK_9);
							robot.keyRelease(KeyEvent.VK_9);
							robot.delay(30);
						}
					}
					robot.keyPress(KeyEvent.VK_PERIOD);
					robot.keyRelease(KeyEvent.VK_PERIOD);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_H);
					robot.keyRelease(KeyEvent.VK_H);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					// Wait about 3 seconds assuming it's ready to save
					robot.delay(2500);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					robot.delay(2000);

					// Checking here if the Current Windows tile is found
					byte[] windowText = new byte[512];
					PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
					robot.delay(2500);
					// Wait until Markup Validation of upload appears
					if (Native.toString(windowText).contains("Confirm Save As")) {
						// Click on Yes button (FIRFOX)
						// Overwrite file
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_Y);
						robot.delay(200);
						robot.keyRelease(KeyEvent.VK_Y);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.delay(5000);
					}

					// Close Source window
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_W);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_W);
					robot.delay(8000);
				} else {
					if (AdditionalInfo.contains("Skip") == true) {
						robot.delay(1000);
					} else {
						if (AdditionalInfo.contains("[Additional Information]") == true) {
							// Launch firefox (Make sure that only 1 tab exists)
							LaunchFirefox.main(args);
							// System.out.println(LaunchFirefox.getPagetitle());
							robot.mouseMove(600, 90);
							robot.mousePress(InputEvent.BUTTON1_MASK);
							robot.mouseRelease(InputEvent.BUTTON1_MASK);
							robot.delay(300);
							robot.keyPress(KeyEvent.VK_DELETE);
							robot.keyRelease(KeyEvent.VK_DELETE);
							robot.delay(1000);

							// Copy URL in the Clipboard
							Toolkit toolkit1 = Toolkit.getDefaultToolkit();
							Clipboard clipboard1 = toolkit1
									.getSystemClipboard();
							StringSelection strSel1 = new StringSelection(
									cell.getContents());
							clipboard1.setContents(strSel1, null);

							// Press Ctrl+V to enter url
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_V);
							robot.delay(200);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_V);
							robot.delay(500);
							// Press Reload current page
							robot.mouseMove(840, 90);
							robot.keyPress(KeyEvent.VK_ENTER);
							robot.keyRelease(KeyEvent.VK_ENTER);
							// Wait 8 seconds assuming that the page is loaded
							robot.delay(8000);

							// Press View Source from Firefox
							robot.keyPress(KeyEvent.VK_ALT);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_U);
							robot.delay(300);
							robot.keyRelease(KeyEvent.VK_ALT);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.keyRelease(KeyEvent.VK_U);

							// Press Ctrl+S to save the html file & press save
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_S);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_S);
							robot.delay(2000);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_C);
							robot.keyRelease(KeyEvent.VK_C);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_SEMICOLON);
							robot.delay(30);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.keyRelease(KeyEvent.VK_SEMICOLON);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_BACK_SLASH);
							robot.keyRelease(KeyEvent.VK_BACK_SLASH);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_E);
							robot.keyRelease(KeyEvent.VK_E);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_M);
							robot.keyRelease(KeyEvent.VK_M);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_P);
							robot.keyRelease(KeyEvent.VK_P);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_BACK_SLASH);
							robot.keyRelease(KeyEvent.VK_BACK_SLASH);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_U);
							robot.keyRelease(KeyEvent.VK_U);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_R);
							robot.keyRelease(KeyEvent.VK_R);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SUBTRACT);
							robot.keyRelease(KeyEvent.VK_SUBTRACT);
							robot.delay(30);
							// ///////////////
							String s = Integer.toString(Url_ToBeTest);
							char[] charArr2 = s.toCharArray();
							if (charArr2.length == 2) {
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30); 
							}
							if (charArr2.length == 1) {
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30);
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30);
							}
							for (int i1 = 0; i1 < charArr2.length; i1++) {
								String Data = String.copyValueOf(charArr2, i1,
										1);
								if (Data.contentEquals("0")) {
									robot.keyPress(KeyEvent.VK_0);
									robot.keyRelease(KeyEvent.VK_0);
									robot.delay(30);
								}
								if (Data.contentEquals("1")) {
									robot.keyPress(KeyEvent.VK_1);
									robot.keyRelease(KeyEvent.VK_1);
									robot.delay(30);
								}
								if (Data.contentEquals("2")) {
									robot.keyPress(KeyEvent.VK_2);
									robot.keyRelease(KeyEvent.VK_2);
									robot.delay(30);
								}
								if (Data.contentEquals("3")) {
									robot.keyPress(KeyEvent.VK_3);
									robot.keyRelease(KeyEvent.VK_3);
									robot.delay(30);
								}
								if (Data.contentEquals("4")) {
									robot.keyPress(KeyEvent.VK_4);
									robot.keyRelease(KeyEvent.VK_4);
									robot.delay(30);
								}
								if (Data.contentEquals("5")) {
									robot.keyPress(KeyEvent.VK_5);
									robot.keyRelease(KeyEvent.VK_5);
									robot.delay(30);
								}
								if (Data.contentEquals("6")) {
									robot.keyPress(KeyEvent.VK_6);
									robot.keyRelease(KeyEvent.VK_6);
									robot.delay(30);
								}
								if (Data.contentEquals("7")) {
									robot.keyPress(KeyEvent.VK_7);
									robot.keyRelease(KeyEvent.VK_7);
									robot.delay(30);
								}
								if (Data.contentEquals("8")) {
									robot.keyPress(KeyEvent.VK_8);
									robot.keyRelease(KeyEvent.VK_8);
									robot.delay(30);
								}
								if (Data.contentEquals("9")) {
									robot.keyPress(KeyEvent.VK_9);
									robot.keyRelease(KeyEvent.VK_9);
									robot.delay(30);
								}
							}
							robot.keyPress(KeyEvent.VK_PERIOD);
							robot.keyRelease(KeyEvent.VK_PERIOD);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_H);
							robot.keyRelease(KeyEvent.VK_H);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_T);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_M);
							robot.keyRelease(KeyEvent.VK_M);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_L);
							// Wait about 3 seconds assuming it's ready to save
							// (May have Perfomance issue with Firefox here)
							robot.delay(3000);
							robot.keyPress(KeyEvent.VK_ENTER);
							robot.keyRelease(KeyEvent.VK_ENTER);
							robot.delay(2000);

							// Checking if Current Pop up is found
							byte[] windowText = new byte[512];
							PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
							User32.INSTANCE.GetWindowTextA(hwnd, windowText,
									512);
							robot.delay(3000);
							// Wait until Markup Validation of upload appears
							if (Native.toString(windowText).contains(
									"Confirm Save As")) {
								// Click on Yes button (FIRFOX)
								// Overwrite file
								robot.keyPress(KeyEvent.VK_ALT);
								robot.keyPress(KeyEvent.VK_Y);
								robot.delay(200);
								robot.keyRelease(KeyEvent.VK_Y);
								robot.keyRelease(KeyEvent.VK_ALT);
								robot.delay(5000);
							}

							// Wait 3 second assuming the html file is saved
							robot.delay(3000);

							// Close Source window
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_W);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_W);
							robot.delay(7000);
						}
					}
				}

				// Close CSE if already opened
				Runtime.getRuntime().exec("taskkill /F /IM cse140.exe ");
				robot.delay(3000);
				// Open the CSE HTML Validator Pro Batch Wizard
				Runtime.getRuntime()
						.exec("cmd /c start C:\\\"Program Files (x86)\\HTMLValidator140\\cse140.exe");
				robot.delay(4500);

				// Check window CSE HTML Validator Pro v14.00 -
				byte[] windowText = new byte[512];
				PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
				User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
				if (Native.toString(windowText).contains(
						"CSE HTML Validator Pro") == false) {
					// System.out.println(Native.toString(windowText));
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					robot.delay(500);
					byte[] windowText1 = new byte[512];
					PointerType hwnd1 = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd1, windowText1, 512);
					if (Native.toString(windowText1).contains(
							"CSE HTML Validator Pro") == false) {
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
					}
				} else {
					robot.delay(500);
				}

				// Click on caption of CSE HTML Validator Pro
				robot.mouseMove(500, 10);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(500);

				robot.keyPress(KeyEvent.VK_F2);
				robot.keyRelease(KeyEvent.VK_F2);
				robot.delay(1000);

				// Select Target File
				robot.mouseMove(250, 250);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1000);

				// Enter Target file location
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_SEMICOLON);
				robot.delay(30);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.keyRelease(KeyEvent.VK_SEMICOLON);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_BACK_SLASH);
				robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_E);
				robot.keyRelease(KeyEvent.VK_E);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_P);
				robot.keyRelease(KeyEvent.VK_P);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_BACK_SLASH);
				robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_U);
				robot.keyRelease(KeyEvent.VK_U);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_R);
				robot.keyRelease(KeyEvent.VK_R);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SUBTRACT);
				robot.keyRelease(KeyEvent.VK_SUBTRACT);
				robot.delay(30);
				String s = Integer.toString(Url_ToBeTest);
				char[] charArr2 = s.toCharArray();
				if (charArr2.length == 2) {
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30); 
				}
				if (charArr2.length == 1) {
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30);
				}
				for (int i1 = 0; i1 < charArr2.length; i1++) {
					String Data = String.copyValueOf(charArr2, i1, 1);
					if (Data.contentEquals("0")) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
					}
					if (Data.contentEquals("1")) {
						robot.keyPress(KeyEvent.VK_1);
						robot.keyRelease(KeyEvent.VK_1);
						robot.delay(30);
					}
					if (Data.contentEquals("2")) {
						robot.keyPress(KeyEvent.VK_2);
						robot.keyRelease(KeyEvent.VK_2);
						robot.delay(30);
					}
					if (Data.contentEquals("3")) {
						robot.keyPress(KeyEvent.VK_3);
						robot.keyRelease(KeyEvent.VK_3);
						robot.delay(30);
					}
					if (Data.contentEquals("4")) {
						robot.keyPress(KeyEvent.VK_4);
						robot.keyRelease(KeyEvent.VK_4);
						robot.delay(30);
					}
					if (Data.contentEquals("5")) {
						robot.keyPress(KeyEvent.VK_5);
						robot.keyRelease(KeyEvent.VK_5);
						robot.delay(30);
					}
					if (Data.contentEquals("6")) {
						robot.keyPress(KeyEvent.VK_6);
						robot.keyRelease(KeyEvent.VK_6);
						robot.delay(30);
					}
					if (Data.contentEquals("7")) {
						robot.keyPress(KeyEvent.VK_7);
						robot.keyRelease(KeyEvent.VK_7);
						robot.delay(30);
					}
					if (Data.contentEquals("8")) {
						robot.keyPress(KeyEvent.VK_8);
						robot.keyRelease(KeyEvent.VK_8);
						robot.delay(30);
					}
					if (Data.contentEquals("9")) {
						robot.keyPress(KeyEvent.VK_9);
						robot.keyRelease(KeyEvent.VK_9);
						robot.delay(30);
					}
				}
				robot.keyPress(KeyEvent.VK_PERIOD);
				robot.keyRelease(KeyEvent.VK_PERIOD);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_H);
				robot.keyRelease(KeyEvent.VK_H);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.delay(1000);
				// Wait about 1 seconds assuming it's ready to save
				// ( May have Perfomance issue here with Firefox)

				// Open file
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_O);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_O);
				robot.delay(1500);

				// Click on Startbutton
				robot.mouseMove(400, 110);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1000);

				// Click on Start processing now
				robot.mouseMove(250, 300);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(10000);

				///////////////// MAKE SURE FIREFOX IS YOUR DEFAULT BROWSER
				// Will Close tab(s) from Browser Firefox
				String PageName="";
				PageName=CloseTabs.PageTitle(PageName);
	            //System.out.println("PageName="+PageName);
	            
				// Minimized FIREFOX window
				robot.mouseMove(1180, 10);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1500);

				// Select all Target list from the Validator Pro Batch Wizard
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_W);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_W);
				robot.delay(500);

				// Then don't save target lish
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_N);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_N);
				robot.delay(500);

				// Populate Title Direct Input for CSE
				// robot.delay(4000);
				String contentSource = "";
				// String TitleContent1 = "";
				int startTitle = 0;
				try {
					// Read Source file
					int linenumberSource = 0;
					boolean titleFound = false;
					// Reading WPSS results
					BufferedReader reader = new BufferedReader(
							new InputStreamReader(new FileInputStream(
									"c:\\Temp\\URL-" + Url_Number + ".html"),
									"UTF8"));
					String ReadCurrentLineResult = null;
					StringBuilder stringBuilder = new StringBuilder();
					while ((ReadCurrentLineResult = reader.readLine()) != null) {
						linenumberSource = linenumberSource + 1;
						stringBuilder.append(ReadCurrentLineResult.toString());
						stringBuilder.append("\n");
						//ReadCurrentLineResult = ReadCurrentLineResult.replaceAll("&amp;",
						//				"&");	
						// contentSource = contentSource +
						// ReadCurrentLineResult+"\n";
						if ((ReadCurrentLineResult.contains("<title") == true)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("</title>") == true)) {
							String str1 = ReadCurrentLineResult.toString();
							int PositionStart = str1.indexOf("<title");
							int PositionEnd = str1.lastIndexOf("</title>");
							char[] charArray = str1.toCharArray();
							TitleStripped = String.copyValueOf(charArray,
									PositionStart, PositionEnd - PositionStart);
							TitleStripped = TitleStripped.replaceAll("<title>",
									"");
								TitleStripped = TitleStripped.replaceAll(">", "");
							titleFound = true;
							// System.out.println("TitleStripped="+TitleStripped);
						}
						if ((ReadCurrentLineResult.contains("<title") == true)
								&& (ReadCurrentLineResult.contains("</title>") == false)) {
							startTitle = linenumberSource;
						}
						if ((startTitle != 0)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("<title") == false)
								&& (ReadCurrentLineResult.contains("</title>") == false)) {
							String str1 = ReadCurrentLineResult;
							int PositionStart = 1;
							int PositionEnd = str1.length();
							char[] charArray = str1.toCharArray();
							TitleStripped = String.copyValueOf(charArray,
									PositionStart, PositionEnd - PositionStart);
							// System.out.println("TitleStripped="+TitleStripped);
						}
						if ((startTitle != 0)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("</title>") == true)) {
							String str1 = ReadCurrentLineResult;
							int PositionStart = 0;
							int PositionEnd = str1.indexOf("</title>");
							char[] charArray = str1.toCharArray();

							String Trimmer = String.copyValueOf(charArray,
									PositionStart, PositionEnd);
							Trimmer = Trimmer.replaceAll("	", "");
							TitleStripped = TitleStripped.replaceAll(">", "");
							TitleStripped = TitleStripped + Trimmer;
							// System.out.println("TitleStripped=" +
							// TitleStripped);
							startTitle = 0;
							titleFound = true;
						}
					}

					contentSource = stringBuilder.toString();
					// Copy content in the Clipboard
					Toolkit toolkit3 = Toolkit.getDefaultToolkit();
					Clipboard clipboard3 = toolkit3.getSystemClipboard();
					StringSelection strSel3 = new StringSelection(contentSource);
					clipboard3.setContents(strSel3, null);
					//System.out.println(contentSource);
					// length = stringBuilder.length();

					// System.out.println("TitleStripped=" + TitleStripped);

				}
				// End of Checking Results file
				catch (Exception ex) {
					JOptionPane pane1 = new JOptionPane(
							"Problem occurs extracting/manipuating the Source file\n"
									+ "Error/Warning occurs in line 1356\n\n"
									+ "Please make sure that the script was not interrupted during the process.\n\n"
									+ ex + "\n\n" + "Press Ok to continue");
					JDialog d1 = pane1.createDialog((JFrame) null,
							"Warning - Problem extracting <TITLE> content");
					d1.setLocation(400, 500);
					d1.setVisible(true);
				}

				// ///////////////////////////////////////////////////////////////
				TitleStripped = TitleStripped.trim();
				// Convert meta charset ISO-8859-1 to UTF-8
				TitleStripped = TitleStripped.replaceAll("&#233;","é");
				TitleStripped = TitleStripped.replaceAll("&#192;","À");
				TitleStripped = TitleStripped.replaceAll("&#224;","à");
				TitleStripped = TitleStripped.replaceAll("&#39;","'");
				TitleStripped = TitleStripped.replaceAll("&#232;","è");
			    TitleStripped = TitleStripped.replaceAll("&#171;","«");
				TitleStripped = TitleStripped.replaceAll("&#187;","»");
				TitleStripped = TitleStripped.replaceAll("&#201;","É");
				TitleStripped = TitleStripped.replaceAll("&amp;", "&");	
				//System.out.println("TitleStripped=" + TitleStripped);
				// ///////////////////////////////////////////////////////////////
				// Enter the Title page in the Accessibility tab
				Label l = new Label(Url_ToBeTest, 4, TitleStripped,
						arial9formatBold);
				sheet2.addCell(l);
				// //////////////////////////////////////////////////////////////
				// Enter the Title page in the Interoperability tab
				WritableCellFeatures cellFeatures4 = new WritableCellFeatures();
				Label label4 = new Label(Url_ToBeTest, 4, TitleStripped,
						arial9formatBold);
				label4.setCellFeatures(cellFeatures4);
				sheet4.addCell(label4);
				// //////////////////////////////////////////////////////////////
				if (((TitleStripped.contains("-") == false && TitleStripped.contains("|") == false)
						|| TitleStripped.length() < 30 == true)) {
					WCAG2_Error2_4_2 = true;	
					
				}

				// Copy the URL address in the Interoperability tab
				Label label5 = new Label(Url_ToBeTest, 5, cell.getContents(),
						arial9formatBold);
				label5.setCellFeatures(cellFeatures4);
				sheet4.addCell(label5);

				// Close CSE
				Runtime.getRuntime().exec("taskkill /F /IM cse140.exe ");
				robot.delay(1500);
			}

			if (Tool.contains("CSE") && (SourceCode.contains("Url"))) {
				// Close CSE if already opened
				Runtime.getRuntime().exec("taskkill /F /IM cse140.exe ");
				robot.delay(3000);
				// Open the CSE HTML Validator Pro Batch Wizard
				Runtime.getRuntime()
						.exec("cmd /c start C:\\\"Program Files (x86)\\HTMLValidator140\\cse140.exe");
				robot.delay(4500);

				// Copy URL content in the Clipboard
				Toolkit toolkit = Toolkit.getDefaultToolkit();
				Clipboard clipboard = toolkit.getSystemClipboard();
				StringSelection strSel = new StringSelection(cell.getContents());
				clipboard.setContents(strSel, null);

				// Check window CSE HTML Validator Pro v14.00
				byte[] windowText = new byte[512];
				PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
				User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);

				if (Native.toString(windowText).contains(
						"CSE HTML Validator Pro") == false) {
					// System.out.println(Native.toString(windowText));
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					robot.delay(500);
					byte[] windowText1 = new byte[512];
					PointerType hwnd1 = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd1, windowText1, 512);
					if (Native.toString(windowText1).contains(
							"CSE HTML Validator Pro") == false) {
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
					}
				} else {
					robot.delay(500);
				}

				// Click on caption of CSE HTML Validator Pro
				robot.mouseMove(500, 10);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(500);

				robot.keyPress(KeyEvent.VK_F2);
				robot.keyRelease(KeyEvent.VK_F2);
				robot.delay(1000);

				// Select URL Target
				robot.mouseMove(250, 320);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(500);

				// Delete old url
				robot.keyPress(KeyEvent.VK_DELETE);
				robot.keyRelease(KeyEvent.VK_DELETE);
				robot.delay(500);

				// Paste new URL
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_V);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_V);
				robot.delay(1000);

				// Continue button does not appear the same position
				// Must tab 5 time instead
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(300);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(300);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(300);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(300);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(300);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				robot.delay(300);

				// Press Add Target
				robot.mouseMove(500, 370);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(500);

				// Click on Startbutton
				robot.mouseMove(400, 110);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1000);

				// Click on Start processing now
				robot.mouseMove(250, 300);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(10000);

				// /////////////// MAKE SURE FIREFOX IS YOUR DEFAULT BROWSER
				// Will Close tab(s) from Browser Firefox
				String PageName="";
				PageName=CloseTabs.PageTitle(PageName);
	            //System.out.println("PageName="+PageName);
				
				
				// Minimized FIREFOX window
				robot.mouseMove(1180, 10);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1500);

				// Select all Target list from the Validator Pro Batch Wizard
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_W);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_W);
				robot.delay(500);

				// Then don't save target lish
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_N);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_N);
				robot.delay(500);

				// Close CSE
				Runtime.getRuntime().exec("taskkill /F /IM cse140.exe ");
				robot.delay(1500);
			}

			if (Tool.contains("CSE")) {
				for (int j = 0; j < 38; j++) {
					try {
						// Read CSE results & enter it to the spreadsheet
						FileInputStream inputFile = new FileInputStream(path3);
						InputStreamReader fr = new InputStreamReader(inputFile,
								"UTF-8");
						// FileReader fr = new FileReader(path3);
						BufferedReader br = new BufferedReader(fr);
						String ReadCurrentLine;
						String MsgDesc1Line1 = "";
						String MsgDesc1Line2 = "";
						String MsgDesc2Line1 = "";
						String MsgDesc2Line2 = "";
						String MsgDesc3Line1 = "";
						String MsgDesc3Line2 = "";
						String MsgDesc4Line1 = "";
						String MsgDesc4Line2 = "";
						String MsgDesc5Line1 = "";
						String MsgDesc5Line2 = "";
						String MsgDesc = "";
						String MsgDesc1 = "";
						String MsgDesc2 = "";
						String MsgDesc3 = "";
						String MsgDesc4 = "";
						String MsgDesc5 = "";
						String CommentDesc1 = "";
						String CommentDesc2 = "";
						String CommentDesc3 = "";
						String CommentDesc4 = "";
						String CommentDesc5 = "";
						int linenumber = 0;
						int commentLine = 0;
						int NumOfInstance1 = 0;
						int NumOfInstance2 = 0;
						int NumOfInstance3 = 0;
						int NumOfInstance4 = 0;
						int NumOfInstance5 = 0;
						String ErrorLocation = "";
						String SuffTech = "";
						while ((ReadCurrentLine = br.readLine()) != null) {
							linenumber = linenumber + 1;
							if ((ReadCurrentLine.contains(WCAG2StringArray[j]) == true)
									// && (WCAG2StringArray[j] == "1.3.1")
									&& (ReadCurrentLine
											.contains("Accessibility Error") == true)) {
								
								String str1 = ReadCurrentLine;
								// System.out.println("str1=" + str1);
								int EndOfLine = str1
										.indexOf("<br>&nbsp;</p></td></tr>");
								int Position = str1.indexOf(" [A");
								int PositionLast = str1.lastIndexOf("]");
								int PositionLast2 = str1.indexOf("].");
								int PositionVisit = str1.indexOf("Visit <a");
								int MsgPos = str1.lastIndexOf("msgdescription");
								int LocPos = str1.lastIndexOf("msgtypecell");
								char[] charArray = str1.toCharArray();
								if ((Position < PositionLast2)
										&& PositionLast2 < PositionLast) {
									SuffTech = String.copyValueOf(charArray,
											Position, PositionLast2 - Position
													+ 1);
								} else {
									SuffTech = String.copyValueOf(charArray,
											Position, PositionLast - Position
													+ 1);
								}
								if ((MsgPos <= Position)
										&& (Position <= PositionVisit)) {
									MsgDesc = String.copyValueOf(charArray,
											MsgPos + 16, Position - MsgPos);
								} else {
									if (Position > PositionVisit) {
										//System.out.println("Pos>PosVisit");
										if (PositionVisit > 0) {
											MsgDesc = String
													.copyValueOf(
															charArray,
															MsgPos + 16,
															Position
																	- MsgPos
																	- (Position - PositionVisit)
																	- 16);
										} else {
											// Exception appears here about 1.3.1 reporting 
											MsgDesc = String.copyValueOf(
													charArray, MsgPos,
													EndOfLine - MsgPos);
											MsgDesc = MsgDesc.replaceAll("</span><p class=\"docsource htmlsource\">",
													"");
											
										}
									} else {
										if (Position >= MsgPos) {
											MsgDesc = String.copyValueOf(
													charArray, MsgPos + 16,
													Position - MsgPos - 17);
											// System.out.println("Position "+Position
											// +">MsgPos"+MsgPos);
										}
									}
								}
								ErrorLocation = String.copyValueOf(charArray,
										33, LocPos - 49);
								// System.out.println("SuffTech=" + SuffTech
								// + " ErrorLocation=" + ErrorLocation);
								if (MsgDesc1 == "" == true) {
									MsgDesc1 = SuffTech.trim();
									MsgDesc1Line2 = MsgDesc.trim();
									// System.out.println("MsgDesc1>"+MsgDesc1);
									// System.out.println("MsgDesc1Line2="+MsgDesc);
								} else {
									if ((MsgDesc2 == "")
											&& (MsgDesc1
													.toString()
													.contentEquals(
															SuffTech.trim()
																	.toString()) == false)) {
										MsgDesc2 = SuffTech.trim();
										MsgDesc2Line2 = MsgDesc.trim();
										// System.out.println("MsgDesc2="+MsgDesc2);
										// System.out.println("MsgDesc2Line2="+MsgDesc);
									} else {
										if ((MsgDesc3 == "")
												&& (MsgDesc1
														.toString()
														.contentEquals(
																SuffTech.trim()
																		.toString()) == false)
												&& (MsgDesc2
														.toString()
														.contentEquals(
																SuffTech.trim()
																		.toString()) == false)) {
											MsgDesc3 = SuffTech.trim();
											MsgDesc3Line2 = MsgDesc.trim();
											// System.out.println("MsgDesc3="+MsgDesc3);
											// System.out.println("MsgDesc3Line2="+MsgDesc);
										} else {
											if ((MsgDesc4 == "")
													&& ((MsgDesc1
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false))
													&& (MsgDesc2
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false)
													&& (MsgDesc3
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false)) {
												MsgDesc4 = SuffTech.trim();
												MsgDesc4Line2 = MsgDesc.trim();
												// System.out.println("MsgDesc4="+MsgDesc4);
												// System.out.println("MsgDesc4Line2="+MsgDesc);
											} else {
												if ((MsgDesc5 == "")
														&& ((MsgDesc1
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false))
														&& (MsgDesc2
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)
														&& (MsgDesc3
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)
														&& (MsgDesc4
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)) {
													MsgDesc5 = SuffTech.trim();
													MsgDesc5Line2 = MsgDesc
															.trim();
													// System.out.println("MsgDesc5="+MsgDesc5);
													// System.out.println("MsgDesc5Line2="+MsgDesc);
												}
											}
										}
									}
								}
								if ((MsgDesc1.toString().contentEquals(
										SuffTech.toString().trim()) == true)
										&& (ReadCurrentLine
												.contains("Accessibility Error") == true)
										&& (MsgDesc1.toString().contentEquals(
												"") == false)) {
									// System.out.println(MsgDesc1);
									MsgDesc1 = SuffTech.toString().trim();
									MsgDesc1Line2 = MsgDesc.toString().trim();
									MsgDesc1Line1 = MsgDesc1Line1.toString()
											+ ErrorLocation + "; ";
									NumOfInstance1 = NumOfInstance1 + 1;
								} else {
									if ((MsgDesc2.toString().contentEquals(
											SuffTech.toString().trim()) == true)
											&& (MsgDesc1
													.toString()
													.contentEquals(
															SuffTech.trim()
																	.toString()) == false)) {
										MsgDesc2 = SuffTech.toString().trim();
										MsgDesc2Line2 = MsgDesc.toString().trim();
										MsgDesc2Line1 = MsgDesc2Line1.toString()
												+ ErrorLocation + "; ";
										NumOfInstance2 = NumOfInstance2 + 1;
									} else {
										if ((MsgDesc3.toString().contentEquals(
												SuffTech.toString().trim()) == true)
												&& (MsgDesc1
														.toString()
														.contentEquals(
																SuffTech.trim()
																		.toString()) == false)
												&& (MsgDesc2
														.toString()
														.contentEquals(
																SuffTech.trim()
																		.toString()) == false)) {
											MsgDesc3 = SuffTech.toString().trim();
											MsgDesc3Line2 = MsgDesc.toString().trim();
											MsgDesc3Line1 = MsgDesc3Line1
													+ ErrorLocation + "; ";
											NumOfInstance3 = NumOfInstance3 + 1;
										} else {
											if ((MsgDesc4
													.toString()
													.contentEquals(
															SuffTech.trim()
																	.toString()) == true)
													&& ((MsgDesc1
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false))
													&& (MsgDesc2
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false)
													&& (MsgDesc3
															.toString()
															.contentEquals(
																	SuffTech.trim()
																			.toString()) == false)) {
												MsgDesc4 = SuffTech.trim();
												MsgDesc4Line2 = MsgDesc.trim();
												MsgDesc4Line1 = MsgDesc4Line1
														+ ErrorLocation + "; ";
												NumOfInstance4 = NumOfInstance4 + 1;
											} else {
												if ((MsgDesc5
														.toString()
														.contentEquals(
																SuffTech.trim()
																		.toString()) == true)
														&& ((MsgDesc1
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false))
														&& (MsgDesc2
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)
														&& (MsgDesc3
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)
														&& (MsgDesc4
																.toString()
																.contentEquals(
																		SuffTech.trim()
																				.toString()) == false)) {
													MsgDesc5 = SuffTech.trim();
													MsgDesc5Line2 = MsgDesc
															.trim();
													MsgDesc5Line1 = MsgDesc5Line1
															+ ErrorLocation
															+ "; ";
													NumOfInstance5 = NumOfInstance5 + 1;
												}
											}
										}
									}
								}
							}
						}
						if (MsgDesc1 == "" == false) {
							if (MsgDesc1Line2.contains("Fails validation,")) {
								CommentDesc1 = MsgDesc1
										+ "\n"
										+ "    Each web page should have no error from the W3C Markup validaton Service"
										+ "\n"
										+ "    Validate your source code at http://validator.w3.org/#validate-by-input"
										+ "\n\n";
								commentLine = 10;
							} else {
								CommentDesc1 = MsgDesc1
										+ " "
										+ MsgDesc1Line2
										+ "\n"
										+ "    Found in Source Line:Column "
										+ MsgDesc1Line1
										+ "\n"
										+ "    Number of Instance: "
										+ NumOfInstance1
										+ " found in CSE HTML Validator PRO v14.0"
										+ "\n\n";
								commentLine = 15;
							}
						}
						if (MsgDesc2 == "" == false) {
							CommentDesc2 = MsgDesc2 + " " + MsgDesc2Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc2Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance2
									+ " found in CSE HTML Validator PRO v14.0"
									+ "\n\n";
							commentLine = 23;
						}
						if (MsgDesc3 == "" == false) {
							CommentDesc3 = MsgDesc3 + " " + MsgDesc3Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc3Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance3
									+ " found in CSE HTML Validator PRO v14.0"
									+ "\n\n";
							commentLine = 33;
						}
						if (MsgDesc4 == "" == false) {
							CommentDesc4 = MsgDesc4 + " " + MsgDesc4Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc4Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance4
									+ " found in CSE HTML Validator PRO v14.0"
									+ "\n\n";
							commentLine = 43;
						}
						if (MsgDesc5 == "" == false) {
							CommentDesc5 = MsgDesc5 + " " + MsgDesc5Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc5Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance5
									+ " found in CSE HTML Validator PRO v14.0";
							commentLine = 47;
						}
						WritableCellFeatures cellFeatures = new WritableCellFeatures();
						cellFeatures.setComment(CommentDesc1 + CommentDesc2
								+ CommentDesc3 + CommentDesc4 + CommentDesc5,
								7, commentLine);
						Label label = new Label(Url_ToBeTest, WCAG2RowArray[j],
								"Fail", arial9formatNoBold);
						label.setCellFeatures(cellFeatures);
						sheet2.addCell(label);
						// End of Read Result file from WPSS tool

						// The remaining results will pass or N/A
						if (MsgDesc1 == "" == true) {
							if (WCAG2StringArray[j] == "1.2.1"
									|| WCAG2StringArray[j] == "1.2.2"
									|| WCAG2StringArray[j] == "1.2.3"
									|| WCAG2StringArray[j] == "1.2.4"
									|| WCAG2StringArray[j] == "1.2.5"
									|| WCAG2StringArray[j] == "3.3.4") {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										WCAG2RowArray[j], "N/A",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							} else {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										WCAG2RowArray[j], "Pass",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							}
						}
					} // End of all URL While procedures
					catch (Exception ex) {
						// All cells modified/added. Now write out the workbook
						JOptionPane pane11 = new JOptionPane(
								"Problem occurs writing CSE result in Spreadsheet\n\n"
										+ "Error/Warning occurs in line 1949\n\n"
										+ "Please make sure that the script was not interrupted during the process.\n"
										+ ex + "\n\n" + "Press Ok to continue");
						JDialog d11 = pane11.createDialog((JFrame) null,
								"Probem writing result in Spreadsheet file");
						d11.setLocation(400, 500);
						d11.setVisible(true);
						// copy.write();
						// copy.close();
						// System.exit(0);
					}
				}
			}
			int length = 0;
			if ((Tool.contains("WPSS") || Tool.contains("W3C"))
					&& SourceCode.contains("Direct Input")) {
				robot.delay(1500);
				if (Tool.contains("WPSS")) {
					// Close PWGSC WPSS tool if already opened
					Runtime.getRuntime().exec("taskkill /F /IM perl.exe ");
					robot.delay(10000);
					// Open the PWGSC WPSS tool
					// Runtime.getRuntime()
					// .exec("cmd /c start C:\\\"Program Files\\WPSS_Tool\\wpss_tool.pl");
				}

				// Prompt Tester you require to follow manual steps
				// (if a sequence of steps need to be done to get to this page)
				if ((AdditionalInfo.contains("[Additional Information]") == false)
						&& (AdditionalInfo.contains("Skip") == false)) {
					LaunchFirefox.main(args);
					JOptionPane pane = new JOptionPane(
							"\n"
									+ "Following instruction need to be done to get to this page\n\n"
									+ AdditionalInfo
									+ "\n\n"
									+ "Press Ok to continue when manual steps are completed");
					JDialog d = pane.createDialog((JFrame) null,
							"Manual steps is required");
					d.setLocation(400, 500);
					d.setVisible(true);

					// Press View Source from Firefox
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_U);
					robot.delay(300);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_U);
					robot.delay(10000);

					// // View Source window appear
					// JOptionPane pane = new JOptionPane("\n"
					// + "Wait until the View Source window appears\n\n"
					// + "then Press Ok to continue");
					// JDialog d = pane.createDialog((JFrame) null,
					// "Wait until the View Source window appears");
					// d.setLocation(400, 500);
					// d.setVisible(true);

					// Press Ctrl+S to save the html file and press save
					// button
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_S);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_S);
					robot.delay(3000);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_SEMICOLON);
					robot.delay(25);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_SEMICOLON);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_BACK_SLASH);
					robot.keyRelease(KeyEvent.VK_BACK_SLASH);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_E);
					robot.keyRelease(KeyEvent.VK_E);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_P);
					robot.keyRelease(KeyEvent.VK_P);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_BACK_SLASH);
					robot.keyRelease(KeyEvent.VK_BACK_SLASH);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_R);
					robot.keyRelease(KeyEvent.VK_R);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_SUBTRACT);
					robot.keyRelease(KeyEvent.VK_SUBTRACT);
					robot.delay(25);
					// ///////////////
					String s = Integer.toString(Url_ToBeTest);
					char[] charArr2 = s.toCharArray();
					if (charArr2.length == 2) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30); 
					}
					if (charArr2.length == 1) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
					}
					for (int i1 = 0; i1 < charArr2.length; i1++) {
						String Data = String.copyValueOf(charArr2, i1, 1);
						if (Data.contentEquals("0")) {
							robot.keyPress(KeyEvent.VK_0);
							robot.keyRelease(KeyEvent.VK_0);
							robot.delay(25);
						}
						if (Data.contentEquals("1")) {
							robot.keyPress(KeyEvent.VK_1);
							robot.keyRelease(KeyEvent.VK_1);
							robot.delay(25);
						}
						if (Data.contentEquals("2")) {
							robot.keyPress(KeyEvent.VK_2);
							robot.keyRelease(KeyEvent.VK_2);
							robot.delay(25);
						}
						if (Data.contentEquals("3")) {
							robot.keyPress(KeyEvent.VK_3);
							robot.keyRelease(KeyEvent.VK_3);
							robot.delay(25);
						}
						if (Data.contentEquals("4")) {
							robot.keyPress(KeyEvent.VK_4);
							robot.keyRelease(KeyEvent.VK_4);
							robot.delay(25);
						}
						if (Data.contentEquals("5")) {
							robot.keyPress(KeyEvent.VK_5);
							robot.keyRelease(KeyEvent.VK_5);
							robot.delay(25);
						}
						if (Data.contentEquals("6")) {
							robot.keyPress(KeyEvent.VK_6);
							robot.keyRelease(KeyEvent.VK_6);
							robot.delay(25);
						}
						if (Data.contentEquals("7")) {
							robot.keyPress(KeyEvent.VK_7);
							robot.keyRelease(KeyEvent.VK_7);
							robot.delay(25);
						}
						if (Data.contentEquals("8")) {
							robot.keyPress(KeyEvent.VK_8);
							robot.keyRelease(KeyEvent.VK_8);
							robot.delay(25);
						}
						if (Data.contentEquals("9")) {
							robot.keyPress(KeyEvent.VK_9);
							robot.keyRelease(KeyEvent.VK_9);
							robot.delay(25);
						}
					}
					// ///////////////
					robot.keyPress(KeyEvent.VK_PERIOD);
					robot.keyRelease(KeyEvent.VK_PERIOD);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_H);
					robot.keyRelease(KeyEvent.VK_H);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(25);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					// Wait about 3 seconds assuming it's ready to save
					// May have a performance issue here if the source file is too big
					robot.delay(3000);
					JOptionPane pane11 = new JOptionPane(
								      "Please make sure that the script does not freeze during the process.\n"
									+ "\n\n" + "Press Ok to continue");
					JDialog d11 = pane11.createDialog((JFrame) null,
							"Alt the program here");
					d11.setLocation(400, 500);
					d11.setVisible(true);
									
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					robot.delay(2000);
					// ///////////////////////////////////

					// Checking here if the Current Windows tile is found
					byte[] windowText = new byte[512];
					PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
					robot.delay(3000);
					// Wait until Markup Validation of upload appears
					if (Native.toString(windowText).contains("Confirm Save As")) {
						// Click on Yes button (FIRFOX)
						// Overwrite file
						robot.keyPress(KeyEvent.VK_ALT);
						robot.keyPress(KeyEvent.VK_Y);
						robot.delay(500);
						robot.keyRelease(KeyEvent.VK_Y);
						robot.keyRelease(KeyEvent.VK_ALT);
						robot.delay(5000);
					}

					// Close Source window
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_W);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_W);
					// View Source window appear

					// JOptionPane pane1 = new JOptionPane("\n"
					// + "Wait until the Browser page appears\n\n"
					// + "then Press Ok to continue");
					// JDialog d1 = pane1.createDialog((JFrame) null,
					// "Wait until the Browser page appears");
					// d1.setLocation(400, 500);
					// d1.setVisible(true);
					robot.delay(8000);
				} else {
					if (AdditionalInfo.contains("Skip") == true) {
						robot.delay(1000);
					} else {
						if (AdditionalInfo.contains("[Additional Information]") == true) {
							// Go to firefox (Home page Environment Canade)
							LaunchFirefox.main(args);
							// System.out.println(LaunchFirefox.getPagetitle());
							robot.mouseMove(600, 90);
							robot.mousePress(InputEvent.BUTTON1_MASK);
							robot.mouseRelease(InputEvent.BUTTON1_MASK);
							robot.delay(300);
							robot.keyPress(KeyEvent.VK_DELETE);
							robot.keyRelease(KeyEvent.VK_DELETE);
							robot.delay(1000);

							// Copy URL in the Clipboard
							Toolkit toolkit = Toolkit.getDefaultToolkit();
							Clipboard clipboard = toolkit.getSystemClipboard();
							StringSelection strSel = new StringSelection(
									cell.getContents());
							clipboard.setContents(strSel, null);

							// Press Ctrl+V to enter url
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_V);
							robot.delay(100);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_V);
							robot.delay(500);
							// Press Reload current page
							robot.mouseMove(840, 90);
							robot.keyPress(KeyEvent.VK_ENTER);
							robot.keyRelease(KeyEvent.VK_ENTER);
							// Wait 8 seconds assuming that the page is loaded
							robot.delay(8000);

							// Press View Source from Firefox
							robot.keyPress(KeyEvent.VK_ALT);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_U);
							robot.delay(100);
							robot.keyRelease(KeyEvent.VK_ALT);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.keyRelease(KeyEvent.VK_U);
							robot.delay(10000);
							// View Source window appear
							// JOptionPane pane = new JOptionPane("\n"
							// + "Wait until the View Source window appears\n\n"
							// + "then Press Ok to continue");
							// JDialog d = pane.createDialog((JFrame) null,
							// "Wait until the View Source window appears");
							// d.setLocation(400, 500);
							// d.setVisible(true);
							// Press Ctrl+S to save the html file & press save
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_S);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_S);
							robot.delay(2000);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_C);
							robot.keyRelease(KeyEvent.VK_C);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_SEMICOLON);
							robot.delay(25);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.keyRelease(KeyEvent.VK_SEMICOLON);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_BACK_SLASH);
							robot.keyRelease(KeyEvent.VK_BACK_SLASH);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_E);
							robot.keyRelease(KeyEvent.VK_E);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_M);
							robot.keyRelease(KeyEvent.VK_M);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_P);
							robot.keyRelease(KeyEvent.VK_P);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_BACK_SLASH);
							robot.keyRelease(KeyEvent.VK_BACK_SLASH);
							robot.delay(25);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_U);
							robot.keyRelease(KeyEvent.VK_U);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_R);
							robot.keyRelease(KeyEvent.VK_R);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SHIFT);
							robot.keyPress(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_SHIFT);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_SUBTRACT);
							robot.keyRelease(KeyEvent.VK_SUBTRACT);
							robot.delay(30);
							// ///////////////
							String s = Integer.toString(Url_ToBeTest);
							char[] charArr2 = s.toCharArray();
							if (charArr2.length == 2) {
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30); 
							}
							if (charArr2.length == 1) {
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30);
								robot.keyPress(KeyEvent.VK_0);
								robot.keyRelease(KeyEvent.VK_0);
								robot.delay(30);
							}
							for (int i1 = 0; i1 < charArr2.length; i1++) {
								String Data = String.copyValueOf(charArr2, i1,
										1);
								if (Data.contentEquals("0")) {
									robot.keyPress(KeyEvent.VK_0);
									robot.keyRelease(KeyEvent.VK_0);
									robot.delay(30);
								}
								if (Data.contentEquals("1")) {
									robot.keyPress(KeyEvent.VK_1);
									robot.keyRelease(KeyEvent.VK_1);
									robot.delay(30);
								}
								if (Data.contentEquals("2")) {
									robot.keyPress(KeyEvent.VK_2);
									robot.keyRelease(KeyEvent.VK_2);
									robot.delay(30);
								}
								if (Data.contentEquals("3")) {
									robot.keyPress(KeyEvent.VK_3);
									robot.keyRelease(KeyEvent.VK_3);
									robot.delay(30);
								}
								if (Data.contentEquals("4")) {
									robot.keyPress(KeyEvent.VK_4);
									robot.keyRelease(KeyEvent.VK_4);
									robot.delay(30);
								}
								if (Data.contentEquals("5")) {
									robot.keyPress(KeyEvent.VK_5);
									robot.keyRelease(KeyEvent.VK_5);
									robot.delay(30);
								}
								if (Data.contentEquals("6")) {
									robot.keyPress(KeyEvent.VK_6);
									robot.keyRelease(KeyEvent.VK_6);
									robot.delay(30);
								}
								if (Data.contentEquals("7")) {
									robot.keyPress(KeyEvent.VK_7);
									robot.keyRelease(KeyEvent.VK_7);
									robot.delay(30);
								}
								if (Data.contentEquals("8")) {
									robot.keyPress(KeyEvent.VK_8);
									robot.keyRelease(KeyEvent.VK_8);
									robot.delay(30);
								}
								if (Data.contentEquals("9")) {
									robot.keyPress(KeyEvent.VK_9);
									robot.keyRelease(KeyEvent.VK_9);
									robot.delay(30);
								}
							}
							// ///////////////
							robot.keyPress(KeyEvent.VK_PERIOD);
							robot.keyRelease(KeyEvent.VK_PERIOD);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_H);
							robot.keyRelease(KeyEvent.VK_H);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_T);
							robot.keyRelease(KeyEvent.VK_T);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_M);
							robot.keyRelease(KeyEvent.VK_M);
							robot.delay(30);
							robot.keyPress(KeyEvent.VK_L);
							robot.keyRelease(KeyEvent.VK_L);
							// Wait about 3 seconds assuming it's ready to save
							// (Perfomance issue with Firefox)
							JOptionPane pane11 = new JOptionPane(
								      "Please make sure that the script does not freeze during the process.\n"
									+ "\n\n" + "Press Ok to continue");
							JDialog d11 = pane11.createDialog((JFrame) null,
									"Alt the program here");
							d11.setLocation(400, 500);
							d11.setVisible(true);
							
							// ///////////////////////////////////
							robot.delay(3000);
							robot.keyPress(KeyEvent.VK_ENTER);
							robot.keyRelease(KeyEvent.VK_ENTER);
							robot.delay(2000);
							// ///////////////////////////////////
							// Click Yes to overwrite 
							// Checking here if the Current Windows tile is
							// found
							byte[] windowText = new byte[512];
							PointerType hwnd = User32.INSTANCE
									.GetForegroundWindow();
							User32.INSTANCE.GetWindowTextA(hwnd, windowText,
									512);
							robot.delay(3000);
							// Wait until Markup Validation of upload appears
							if (Native.toString(windowText).contains(
									"Confirm Save As")) {
								// Click on Yes button (FIRFOX)
								// Overwrite file
								robot.keyPress(KeyEvent.VK_ALT);
								robot.keyPress(KeyEvent.VK_Y);
								robot.delay(100);
								robot.keyRelease(KeyEvent.VK_Y);
								robot.keyRelease(KeyEvent.VK_ALT);
							}

							// Wait 3 second assuming the html file is saved
							robot.delay(3000);

							// Close Source window
							robot.keyPress(KeyEvent.VK_CONTROL);
							robot.keyPress(KeyEvent.VK_W);
							robot.keyRelease(KeyEvent.VK_CONTROL);
							robot.keyRelease(KeyEvent.VK_W);
							robot.delay(1000);
						}
					}
				}

				// Populate Title Direct Input for WPSS or W3C
				// /////////////////////////////////////////////////////////////////////////
				// robot.delay(4000);
				String contentSource = "";
				// String TitleContent1 = "";
				int startTitle = 0;
				try {
					// Read Source file
					int linenumberSource = 0;
					boolean titleFound = false;
					// Reading WPSS results
					BufferedReader reader = new BufferedReader(
							new InputStreamReader(new FileInputStream(
									"c:\\Temp\\URL-" + Url_Number + ".html"),
									"UTF8"));
					String ReadCurrentLineResult = null;
					StringBuilder stringBuilder = new StringBuilder();
					while ((ReadCurrentLineResult = reader.readLine()) != null) {
						linenumberSource = linenumberSource + 1;
						stringBuilder.append(ReadCurrentLineResult.toString());
						stringBuilder.append("\n");
						// contentSource = contentSource +
						// ReadCurrentLineResult+"\n";
						if ((ReadCurrentLineResult.contains("<title") == true)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("</title>") == true)) {
							String str1 = ReadCurrentLineResult.toString();
							int PositionStart = str1.indexOf("<title");
							int PositionEnd = str1.lastIndexOf("</title>");
							char[] charArray = str1.toCharArray();
							TitleStripped = String.copyValueOf(charArray,
									PositionStart, PositionEnd - PositionStart);
							TitleStripped = TitleStripped.replaceAll("<title>",
									"");
							TitleStripped = TitleStripped.replaceAll(">", "");
							titleFound = true;
							// System.out.println("TitleStripped="+TitleStripped);
						}
						if ((ReadCurrentLineResult.contains("<title") == true)
								&& (ReadCurrentLineResult.contains("</title>") == false)) {
							startTitle = linenumberSource;
						}
						if ((startTitle != 0)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("<title") == false)
								&& (ReadCurrentLineResult.contains("</title>") == false)) {
							String str1 = ReadCurrentLineResult;
							int PositionStart = 1;
							int PositionEnd = str1.length();
							char[] charArray = str1.toCharArray();
							TitleStripped = String.copyValueOf(charArray,
									PositionStart, PositionEnd - PositionStart);
							// System.out.println("TitleStripped="+TitleStripped);
						}
						if ((startTitle != 0)
								&& (titleFound == false)
								&& (ReadCurrentLineResult.contains("</title>") == true)) {
							String str1 = ReadCurrentLineResult;
							int PositionStart = 0;
							int PositionEnd = str1.indexOf("</title>");
							char[] charArray = str1.toCharArray();

							String Trimmer = String.copyValueOf(charArray,
									PositionStart, PositionEnd);
							Trimmer = Trimmer.replaceAll("	", "");							
							TitleStripped = TitleStripped.replaceAll(">", "");
							TitleStripped = TitleStripped + Trimmer;
							// System.out.println("TitleStripped=" +
							// TitleStripped);
							startTitle = 0;
							titleFound = true;
						}
					}

					contentSource = stringBuilder.toString();
					// Copy content in the Clipboard
					Toolkit toolkit3 = Toolkit.getDefaultToolkit();
					Clipboard clipboard3 = toolkit3.getSystemClipboard();
					StringSelection strSel3 = new StringSelection(contentSource);
					clipboard3.setContents(strSel3, null);
					//System.out.println(contentSource);
					length = stringBuilder.length();

					// System.out.println("TitleStripped=" + TitleStripped);

				}
				// End of Checking Results file
				catch (Exception ex) {
					JOptionPane pane1 = new JOptionPane(
							"Problem occurs extracting/manipuating the Source file\n"
									+ "Error/Warning occurs in line 2522\n\n"
									+ "Please make sure that the script was not interrupted during the process.\n"
									+ ex + "\n\n" + "Press Ok to continue");
					JDialog d1 = pane1.createDialog((JFrame) null,
							"Warning - Problem extracting <TITLE> content");
					d1.setLocation(400, 500);
					d1.setVisible(true);
				}
				// Convert meta charset ISO-8859-1 to UTF-8
				TitleStripped = TitleStripped.replaceAll("&#233;","é");
				TitleStripped = TitleStripped.replaceAll("&#192;","À");
				TitleStripped = TitleStripped.replaceAll("&#224;","à");
				TitleStripped = TitleStripped.replaceAll("&#39;","'");
				TitleStripped = TitleStripped.replaceAll("&#232;","è");
			    TitleStripped = TitleStripped.replaceAll("&#171;","«");
				TitleStripped = TitleStripped.replaceAll("&#187;","»");
				TitleStripped = TitleStripped.replaceAll("&#201;","É");
				TitleStripped = TitleStripped.replaceAll("&amp;", "&");	
				// ///////////////////////////////////////////////////////////////
				// System.out.println("TitleStripped=" + TitleStripped);
				// ///////////////////////////////////////////////////////////////
				// Enter the Title page in the Accessibility tab
				Label l = new Label(Url_ToBeTest, 4, TitleStripped,
						arial9formatBold);
				sheet2.addCell(l);
				// ///////////////////////////////////////////////////////////////
				// Enter the Title page in the Interoperability tab
				WritableCellFeatures cellFeatures4 = new WritableCellFeatures();
				Label label4 = new Label(Url_ToBeTest, 4, TitleStripped,
						arial9formatBold);
				label4.setCellFeatures(cellFeatures4);
				sheet4.addCell(label4);
				// ///////////////////////////////////////////////////////////////
				if (((TitleStripped.contains("-") == false && TitleStripped.contains("|") == false)
						|| TitleStripped.length() < 30 == true)) {
					WCAG2_Error2_4_2 = true;	
					
					// System.out.println("-="+TitleContent.contains("-"));
					// System.out.println("|="+TitleContent.contains("|"));
					// System.out.println("length-="+TitleContent.length());
					// System.out.println("2.4.2 Error found ");
				}

				// Copy the URL address in the Interoperability tab
				Label label5 = new Label(Url_ToBeTest, 5, cell.getContents(),
						arial9formatBold);
				label5.setCellFeatures(cellFeatures4);
				sheet4.addCell(label5);
			}

			if (Tool.contains("WPSS") && SourceCode.contains("Direct Input")) {
				// Direct HTML Input
				// Close PWGSC WPSS tool if already opened
				Runtime.getRuntime().exec("taskkill /F /IM perl.exe ");
				robot.delay(2000);
				// Open the PWGSC WPSS tool
				Runtime.getRuntime()
						.exec("cmd /c start C:\\\"Program Files (x86)\\WPSS_Tool\\wpss_tool.pl");

				robot.delay(3000);
				// get the taskbar's window handle
				HWND shellTrayHwnd1 = User32.instance.FindWindow(
						User32.SHELL_TRAY_WND, null);
				// use it to minimize all windows
				User32.instance.SendMessageA(shellTrayHwnd1, User32.WM_COMMAND,
						User32.MIN_ALL, 0);

				// System.out.println("length="+length);
				if (length < 100000) {
					// Checking Current Windows tile is PWGSC WPSS Validator
					byte[] windowText = new byte[512];
					PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
					robot.delay(3000);
					// User32.INSTANCE.GetForegroundWindow();
					if (Native.toString(windowText).contains(
							"PWGSC WPSS Validator")) {
						// System.out.println("Window "+Native.toString(windowText)+" FOUND");
						// Click on the Direct Input tab
						robot.mouseMove(225, 60);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1000);
						// Click on the Text box
						robot.mouseMove(275, 100);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1000);

						// Press Ctrl+V to enter url
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(800);

						// Click on the Click on the "Check Url List" button
						robot.mouseMove(700, 650);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1500);
					} else {
						// System.out.println("Window "+Native.toString(windowText)+" NOT OPENED YET");
						// Wait another 5 secs if system appears to be too slow
						// TODO Make this a While instruction until PWGSC WPSS
						// Validator appears instead of iF else
						robot.delay(5000);
						// Click on the Direct Input tab
						robot.mouseMove(225, 60);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1000);
						// Click on the Text box
						robot.mouseMove(275, 100);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1000);

						// Press Ctrl+V to enter url
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(1000);

						// Click on the Click on the "Check Url List" button
						robot.mouseMove(700, 650);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1500);
					}
				} else {
					// Read the file instead
					// Click on the URL List tab
					robot.mouseMove(275, 60);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);
					// Click on the Text box
					robot.mouseMove(275, 100);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);

					robot.keyPress(KeyEvent.VK_F);
					robot.keyRelease(KeyEvent.VK_F);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_I);
					robot.keyRelease(KeyEvent.VK_I);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_E);
					robot.keyRelease(KeyEvent.VK_E);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_SEMICOLON);
					robot.delay(30);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_SEMICOLON);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SLASH);
					robot.keyRelease(KeyEvent.VK_SLASH);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SLASH);
					robot.keyRelease(KeyEvent.VK_SLASH);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SLASH);
					robot.keyRelease(KeyEvent.VK_SLASH);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_C);
					robot.keyRelease(KeyEvent.VK_C);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SHIFT);
					robot.keyPress(KeyEvent.VK_SEMICOLON);
					robot.delay(30);
					robot.keyRelease(KeyEvent.VK_SHIFT);
					robot.keyRelease(KeyEvent.VK_SEMICOLON);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SLASH);
					robot.keyRelease(KeyEvent.VK_SLASH);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_E);
					robot.keyRelease(KeyEvent.VK_E);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_P);
					robot.keyRelease(KeyEvent.VK_P);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SLASH);
					robot.keyRelease(KeyEvent.VK_SLASH);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_U);
					robot.keyRelease(KeyEvent.VK_U);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_R);
					robot.keyRelease(KeyEvent.VK_R);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_SUBTRACT);
					robot.keyRelease(KeyEvent.VK_SUBTRACT);
					robot.delay(30);
					// ////////////////
					String s = Integer.toString(Url_ToBeTest);
					char[] charArr2 = s.toCharArray();
					if (charArr2.length == 2) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30); 
					}
					if (charArr2.length == 1) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
					}
					for (int i1 = 0; i1 < charArr2.length; i1++) {
						String Data = String.copyValueOf(charArr2, i1, 1);
						if (Data.contentEquals("0")) {
							robot.keyPress(KeyEvent.VK_0);
							robot.keyRelease(KeyEvent.VK_0);
							robot.delay(30);
						}
						if (Data.contentEquals("1")) {
							robot.keyPress(KeyEvent.VK_1);
							robot.keyRelease(KeyEvent.VK_1);
							robot.delay(30);
						}
						if (Data.contentEquals("2")) {
							robot.keyPress(KeyEvent.VK_2);
							robot.keyRelease(KeyEvent.VK_2);
							robot.delay(30);
						}
						if (Data.contentEquals("3")) {
							robot.keyPress(KeyEvent.VK_3);
							robot.keyRelease(KeyEvent.VK_3);
							robot.delay(30);
						}
						if (Data.contentEquals("4")) {
							robot.keyPress(KeyEvent.VK_4);
							robot.keyRelease(KeyEvent.VK_4);
							robot.delay(30);
						}
						if (Data.contentEquals("5")) {
							robot.keyPress(KeyEvent.VK_5);
							robot.keyRelease(KeyEvent.VK_5);
							robot.delay(30);
						}
						if (Data.contentEquals("6")) {
							robot.keyPress(KeyEvent.VK_6);
							robot.keyRelease(KeyEvent.VK_6);
							robot.delay(30);
						}
						if (Data.contentEquals("7")) {
							robot.keyPress(KeyEvent.VK_7);
							robot.keyRelease(KeyEvent.VK_7);
							robot.delay(30);
						}
						if (Data.contentEquals("8")) {
							robot.keyPress(KeyEvent.VK_8);
							robot.keyRelease(KeyEvent.VK_8);
							robot.delay(30);
						}
						if (Data.contentEquals("9")) {
							robot.keyPress(KeyEvent.VK_9);
							robot.keyRelease(KeyEvent.VK_9);
							robot.delay(30);
						}
					}
					// ///////////////
					robot.keyPress(KeyEvent.VK_PERIOD);
					robot.keyRelease(KeyEvent.VK_PERIOD);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_H);
					robot.keyRelease(KeyEvent.VK_H);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_T);
					robot.keyRelease(KeyEvent.VK_T);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_M);
					robot.keyRelease(KeyEvent.VK_M);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_L);
					robot.keyRelease(KeyEvent.VK_L);

					// Click on the Click on the "Check Url List" button
					robot.mouseMove(700, 650);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1500);
				}
				// Move "Results Window" a bit lower
				robot.mouseMove(275, 5);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseMove(275, 50);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1500);
			}

			if (Tool.contains("W3C")) {
				Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
				robot.delay(1500);
				// Open W3C browser at http://html5.validator.nu
				if (SourceCode.contains("Direct Input")) {
					Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
					robot.delay(2000);
					String[] commands = { "cmd", "/c", "start", "/max",
							"iexplore.exe", "-nohome",
							"http://validator.w3.org/#validate-by-input" };
					Runtime.getRuntime().exec(commands);
					robot.delay(7000);

					// Checking if Current Windows tile is W3C Validator
					byte[] windowText = new byte[512];
					PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
					robot.delay(3000);
					// Wait until IE Browser appears
					if (Native.toString(windowText).contains(
							"Markup Validation Service")) {
						// PageUp
						robot.keyPress(KeyEvent.VK_PAGE_UP);
						robot.keyRelease(KeyEvent.VK_PAGE_UP);

						// Goto the Text box address
						robot.mouseMove(850, 120);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(100);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(300);

						// Tab 6 times to get to the address box
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);

						// Press Ctrl+S to save the html file
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(3500);

						if (length > 50000) {
							robot.delay(3000);
						}
						if (length > 500000) {
							robot.delay(4000);
						}

						// Goto the Check button
						// robot.mouseMove(650, 420);
						// Tab 12 times to get to the Check button
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_UP);
						robot.keyRelease(KeyEvent.VK_UP);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_UP);
						robot.keyRelease(KeyEvent.VK_UP);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						// Press Enter
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						robot.delay(6000);
						if (length > 50000) {
							robot.delay(3000);
						}
						if (length > 500000) {
							robot.delay(4000);
						}
					} else {
						// wait more time to display
						robot.delay(8000);
						// Click on the text box to paste the html code
						robot.mouseMove(650, 275);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(200);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(1000);

						// Press Ctrl+S to save the html file
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(3000);

						// click on the Check button
						robot.mouseMove(650, 440);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(6000);
						if (length > 50000) {
							robot.delay(3000);
						}
						if (length > 500000) {
							robot.delay(4000);
						}
					}
				}

				// Open W3C browser at http://html5.validator.nu
				if (SourceCode.contains("Url")) {
					Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
					robot.delay(2000);
					String[] commands = { "cmd", "/c", "start", "/max",
							"iexplore.exe", "-nohome",
							"http://validator.w3.org/#validate_by_uri" };
					Runtime.getRuntime().exec(commands);
					robot.delay(8000);
					// Checking here if the correct Windows is found
					byte[] windowText = new byte[512];
					PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
					User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
					robot.delay(3000);
					// Wait until IE Browser appears
					if (Native.toString(windowText).contains(
							"Markup Validation Service")) {
						// Copy content in the Clipboard
						Toolkit toolkit = Toolkit.getDefaultToolkit();
						Clipboard clipboard = toolkit.getSystemClipboard();
						StringSelection strSel = new StringSelection(
								cell.getContents());
						clipboard.setContents(strSel, null);

						// PageUp
						robot.keyPress(KeyEvent.VK_PAGE_UP);
						robot.keyRelease(KeyEvent.VK_PAGE_UP);
						robot.delay(300);

						// Goto the Text box address
						robot.mouseMove(850, 120);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(100);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(300);

						// Tab 6 times to get to the address box
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);

						// Press Ctrl+S to save the html file
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(3000);

						// Goto the Check button
						robot.mouseMove(650, 420);
						// Tab 12 times to get to the Check button
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						robot.delay(50);
						// Press Enter
						robot.keyPress(KeyEvent.VK_ENTER);
						robot.keyRelease(KeyEvent.VK_ENTER);
						robot.delay(6000);
						if (length > 50000) {
							robot.delay(3000);
						}
						if (length > 500000) {
							robot.delay(4000);
						}
					} else {
						robot.delay(8000);
						// Copy content in the Clipboard
						Toolkit toolkit = Toolkit.getDefaultToolkit();
						Clipboard clipboard = toolkit.getSystemClipboard();
						StringSelection strSel = new StringSelection(
								cell.getContents());
						clipboard.setContents(strSel, null);

						robot.mouseMove(300, 335);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(300);

						// Press Ctrl+S to save the html file
						robot.keyPress(KeyEvent.VK_CONTROL);
						robot.keyPress(KeyEvent.VK_V);
						robot.keyRelease(KeyEvent.VK_CONTROL);
						robot.keyRelease(KeyEvent.VK_V);
						robot.delay(300);

						// click on the Check button
						robot.mouseMove(650, 440);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.mousePress(InputEvent.BUTTON1_MASK);
						robot.mouseRelease(InputEvent.BUTTON1_MASK);
						robot.delay(6000);
						if (length > 50000) {
							robot.delay(3000);
						}
						if (length > 500000) {
							robot.delay(4000);
						}
					}
				}

				// Right-Click IE
				robot.mouseMove(300, 70);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(600);

				// Select File Menu from IEexplore 8.0
				// robot.mouseMove(20, 70); //IE 8.0 No accessible that
				// robot.mousePress(InputEvent.BUTTON3_MASK);
				// robot.delay(10000);
				// robot.mouseRelease(InputEvent.BUTTON3_MASK);
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_F);
				robot.keyRelease(KeyEvent.VK_F);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.delay(300);

				// robot.mouseMove(30, 213); // IE 8.0
				robot.mouseMove(30, 250); // IE 9.0
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.delay(300);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);

				// Checking here if the Current Windows tile is found
				byte[] windowText = new byte[512];
				PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
				User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
				robot.delay(3000);
				// Wait until Markup Validation of upload appears
				if (Native.toString(windowText).contains(
						"Markup Validation of upload")) {
					// System.out.println("Save Webpage window found"+Native.toString(windowText));
				}
				// TODO Add While command for Checking if Save Webpage exists
				robot.keyPress(KeyEvent.VK_BACK_SPACE);
				robot.keyRelease(KeyEvent.VK_BACK_SPACE);
				// "c:\\Temp\\w3c" + Url_Number + ".txt");
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_SEMICOLON);
				robot.delay(30);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.keyRelease(KeyEvent.VK_SEMICOLON);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_BACK_SLASH);
				robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_E);
				robot.keyRelease(KeyEvent.VK_E);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_P);
				robot.keyRelease(KeyEvent.VK_P);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_BACK_SLASH);
				robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_W);
				robot.keyRelease(KeyEvent.VK_W);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_3);
				robot.keyRelease(KeyEvent.VK_3);
				robot.delay(100);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.delay(30);
				// ///////////////
				String s = Integer.toString(Url_ToBeTest);
				char[] charArr2 = s.toCharArray();
				if (charArr2.length == 2) {
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30); 
				}
				if (charArr2.length == 1) {
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30);
					robot.keyPress(KeyEvent.VK_0);
					robot.keyRelease(KeyEvent.VK_0);
					robot.delay(30);
				}
				for (int i1 = 0; i1 < charArr2.length; i1++) {
					String Data = String.copyValueOf(charArr2, i1, 1);
					if (Data.contentEquals("0")) {
						robot.keyPress(KeyEvent.VK_0);
						robot.keyRelease(KeyEvent.VK_0);
						robot.delay(30);
					}
					if (Data.contentEquals("1")) {
						robot.keyPress(KeyEvent.VK_1);
						robot.keyRelease(KeyEvent.VK_1);
						robot.delay(30);
					}
					if (Data.contentEquals("2")) {
						robot.keyPress(KeyEvent.VK_2);
						robot.keyRelease(KeyEvent.VK_2);
						robot.delay(30);
					}
					if (Data.contentEquals("3")) {
						robot.keyPress(KeyEvent.VK_3);
						robot.keyRelease(KeyEvent.VK_3);
						robot.delay(30);
					}
					if (Data.contentEquals("4")) {
						robot.keyPress(KeyEvent.VK_4);
						robot.keyRelease(KeyEvent.VK_4);
						robot.delay(30);
					}
					if (Data.contentEquals("5")) {
						robot.keyPress(KeyEvent.VK_5);
						robot.keyRelease(KeyEvent.VK_5);
						robot.delay(30);
					}
					if (Data.contentEquals("6")) {
						robot.keyPress(KeyEvent.VK_6);
						robot.keyRelease(KeyEvent.VK_6);
						robot.delay(30);
					}
					if (Data.contentEquals("7")) {
						robot.keyPress(KeyEvent.VK_7);
						robot.keyRelease(KeyEvent.VK_7);
						robot.delay(30);
					}
					if (Data.contentEquals("8")) {
						robot.keyPress(KeyEvent.VK_8);
						robot.keyRelease(KeyEvent.VK_8);
						robot.delay(30);
					}
					if (Data.contentEquals("9")) {
						robot.keyPress(KeyEvent.VK_9);
						robot.keyRelease(KeyEvent.VK_9);
						robot.delay(30);
					}
				}
				// ///////////////
				robot.keyPress(KeyEvent.VK_PERIOD);
				robot.keyRelease(KeyEvent.VK_PERIOD);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_H);
				robot.keyRelease(KeyEvent.VK_H);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.delay(30);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.delay(1000);

				// JOptionPane pane1 = new JOptionPane("\n"
				// + "Check if file is deleted\n\n"
				// + "Press Ok to continue");
				// JDialog d1 = pane1.createDialog((JFrame) null,
				// "Check if file w3c.txt is deleted");
				// d1.setLocation(400, 500);
				// d1.setVisible(true);

				// Window: iexplore.exe: Save Webpage
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_S);
				robot.delay(500);
				robot.keyRelease(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.delay(5000);

				byte[] windowText1 = new byte[512];
				PointerType hwnd1 = User32.INSTANCE.GetForegroundWindow();
				User32.INSTANCE.GetWindowTextA(hwnd1, windowText1, 512);
				robot.delay(500);
				if (Native.toString(windowText1).contains("Confirm Save As")) {
					// Overwrite file (IE 9)
					robot.keyPress(KeyEvent.VK_ALT);
					robot.keyPress(KeyEvent.VK_Y);
					robot.delay(500);
					robot.keyRelease(KeyEvent.VK_Y);
					robot.keyRelease(KeyEvent.VK_ALT);
					robot.delay(2000);
				}

				// Close Internet Explore
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_F4);
				robot.delay(500);
				robot.keyRelease(KeyEvent.VK_F4);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.delay(4000);

				try {
					// Read 4.1.1 results & enter them in spreadsheet
					FileInputStream inputFile = new FileInputStream(
							"c:\\Temp\\w3c" + Url_Number + ".htm");
					InputStreamReader fr4_1_1 = new InputStreamReader(
							inputFile, "UTF-8");
					// FileReader fr4_1_1 = new FileReader("c:\\Temp\\w3c.txt");
					BufferedReader br4_1_1 = new BufferedReader(fr4_1_1);
					// String content4_1_1 = "";
					String ReadCurrentLine4_1_1;
					int linenumber4_1_1 = 0;
					int ErrorFoundLine = 0;
					int WCAG2_4_1_1Row = 54;
					boolean WCAG2_4_1_1Error1 = false;
					String CurrentComment = "";
					while ((ReadCurrentLine4_1_1 = br4_1_1.readLine()) != null)
					// WCAG2.0 4.1.1
					{
						linenumber4_1_1 = linenumber4_1_1 + 1;
						if ((ReadCurrentLine4_1_1.contains("<TD class=") == true)
								&& ReadCurrentLine4_1_1.contains("colSpan=")
								&& ReadCurrentLine4_1_1.contains("invalid")
								&& ReadCurrentLine4_1_1.contains("Error")
								&& WCAG2_4_1_1Error1 == false) {
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("<TD class=", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("colSpan=2", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("invalid", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("\"2", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("\"", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll("</TD></TR>", "");
							ReadCurrentLine4_1_1 = ReadCurrentLine4_1_1
									.replaceAll(">", "");							
							// Failure of Success Criterion WCAG2 4.1.1
							CurrentComment = ("FAILED 4.1.1 Parsing G134: Validating Web the page using the W3C Markup validaton Service from http://validator.w3.org/#validate-by-input\n" + ReadCurrentLine4_1_1);
							WritableCellFeatures cellFeatures4_1_1 = new WritableCellFeatures();
							cellFeatures4_1_1.setComment(CurrentComment, 5, 6);
							Label label4_1_1 = new Label(Url_ToBeTest,
									WCAG2_4_1_1Row, "Fail", arial9formatNoBold);
							label4_1_1.setCellFeatures(cellFeatures4_1_1);
							// System.out.println("Comment="+ReadCurrentLine4_1_1);
							sheet2.addCell(label4_1_1);
							//System.out.println("Errors Found1="+ReadCurrentLine4_1_1);
							WCAG2_4_1_1Error1 = true;
						}
						if ((ReadCurrentLine4_1_1.contains("<td colspan=\"2\" class=\"invalid\">") == true)
								&& WCAG2_4_1_1Error1 == false) {
							// Failure of Success Criterion WCAG2 4.1.1
							//System.out.println("Errors Found2="+ReadCurrentLine4_1_1);
							ErrorFoundLine = linenumber4_1_1 + 1;
							WCAG2_4_1_1Error1 = true;
						}
						if ((WCAG2_4_1_1Error1 == true)
								&& ErrorFoundLine==linenumber4_1_1) {
							CurrentComment = ("FAILED 4.1.1 Parsing G134: Validating Web the page using the W3C Markup validaton Service from http://validator.w3.org/#validate-by-input\n" + ReadCurrentLine4_1_1);
							WritableCellFeatures cellFeatures4_1_1 = new WritableCellFeatures();
							cellFeatures4_1_1.setComment(CurrentComment, 5, 6);
							Label label4_1_1 = new Label(Url_ToBeTest,
									WCAG2_4_1_1Row, "Fail", arial9formatNoBold);
							label4_1_1.setCellFeatures(cellFeatures4_1_1);
							// System.out.println("Comment="+ReadCurrentLine4_1_1);
							sheet2.addCell(label4_1_1);
						} else {
							if ((WCAG2_4_1_1Error1 == false)) {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										54, "Pass",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							}
							
						// end of while procedure for 4.1.1
					}

					// The remaining results will pass or N/A
					for (int j = 0; j < 38; j++) {
						if (WCAG2StringArray[j] != "4.1.1") {
							// Result of Success Criterion WCAG2
							WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
							// cellFeatures.setComment("Failed ",4,2);
							Label label2 = new Label(Url_ToBeTest,
								WCAG2RowArray[j], "N/A",
								arial9formatNoBold);
							label2.setCellFeatures(cellFeatures2);
							sheet2.addCell(label2);
						}
					}
				}
					// End of all URL While procedures
				} catch (Exception ex) {
					// All cells modified/added. Now write out the workbook
					JOptionPane pane1 = new JOptionPane(
							"Problem occurs writing W3C result in Spreadsheet\n\n"
									+ "Error/Warning occurs in line 3449\n\n"
									+ "Please make sure that the script was not interrupted during the process.\n"
									+ ex + "\n\n" + "Press Ok to continue");
					JDialog d1 = pane1.createDialog((JFrame) null,
							"Probem writing result in Spreadsheet file");
					d1.setLocation(400, 500);
					d1.setVisible(true);
				}
			}

			if (Tool.contains("WPSS") && SourceCode.contains("Url")) {
				if (Tool.contains("WPSS")) {
					// Close PWGSC WPSS tool if already opened
					Runtime.getRuntime().exec("taskkill /F /IM perl.exe");
					robot.delay(5000);
					// Open the PWGSC WPSS tool
					Runtime.getRuntime()
							.exec("cmd /c start C:\\\"Program Files (x86)\\WPSS_Tool\\wpss_tool.pl");
					robot.delay(5000);
				}

				// Checking Current Windows is found
				byte[] windowText = new byte[512];
				PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
				User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
				// User32.INSTANCE.GetForegroundWindow();
				robot.delay(3000);
				if (Native.toString(windowText)
						.contains("PWGSC WPSS Validator")) {
					// System.out.println("Window "+Native.toString(windowText)+" FOUND");
					// Click on the Url List tab
					robot.mouseMove(275, 60);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);
					// Click on the Text box
					robot.mouseMove(275, 100);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);
				} else {
					// System.out.println("Window "+Native.toString(windowText)+" NOT OPENED YET");
					// Wait another 5 secs if the system is too slow
					robot.delay(5000);
					// Click on the Url List tab
					robot.mouseMove(275, 60);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);
					// Click on the Text box
					robot.mouseMove(275, 100);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1000);
				}

				// Copy URL content in the Clipboard
				Toolkit toolkit = Toolkit.getDefaultToolkit();
				Clipboard clipboard = toolkit.getSystemClipboard();
				StringSelection strSel = new StringSelection(cell.getContents());
				clipboard.setContents(strSel, null);
				robot.delay(800);

				// Press Ctrl+V to enter url
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_V);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_V);
				robot.delay(800);

				// Click on the Click on the "Check Url List" button
				robot.mouseMove(700, 650);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1500);

				// Move "Results Window" a bit lower
				robot.mouseMove(275, 5);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseMove(275, 50);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1500);
			}

			AnalysisCompleted = false;
			// Save the result every 3 seconds until the analysis is completed
			if (Tool.contains("WPSS")) {
				while (AnalysisCompleted == false) {
					// Bring in front "PWGSC WPSS Validator Window"
					robot.mouseMove(270, 5);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(500);

					// Move "PWGSC WPSS Validator Window" lower
					robot.mouseMove(275, 5);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseMove(275, 200);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(1500);

					// From the Results Window
					// Click on the "ACC" tabg
					robot.mouseMove(270, 100);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);

					// Copy WorkingStorage address in the Clipboard
					Toolkit toolkit = Toolkit.getDefaultToolkit();
					Clipboard clipboard = toolkit.getSystemClipboard();
					StringSelection strSel = new StringSelection(WorkingStorage);
					clipboard.setContents(strSel, null);
					robot.delay(4000);

					// Click on File Menu
					robot.mouseMove(25, 85);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(800);
					// Click on Save As
					robot.mouseMove(55, 100);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(2000);
					// Paste the Working storage location
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);
					robot.delay(1000);

					// Click on the Save button
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);

					//String ContentResult = "";
					// Wait 3 secs assuming the file successfully saved
					robot.delay(5000);

					// ///////////////////////////////////////////////////////////////
					// System.out.println("URL=" + TitleStripped);
					// ///////////////////////////////////////////////////////////////
					// // Enter the Title page in the Accessibility tab
					// Label l = new Label(Url_ToBeTest, 5, cell.getContents(),
					// arial9formatBold);
					// sheet2.addCell(l);
					// //
					// //////////////////////////////////////////////////////////////
					// Enter the URL address in the Interoperability tab
					WritableCellFeatures cellFeatures4 = new WritableCellFeatures();
					Label label4 = new Label(Url_ToBeTest, 5,
							cell.getContents(), arial9formatBold);
					label4.setCellFeatures(cellFeatures4);
					sheet4.addCell(label4);
					// //////////////////////////////////////////////////////////////
					//StringBuilder stringBuilder = new StringBuilder();
									// Read Source file
					BufferedReader inputStream = null;

			        try {
			            inputStream = new BufferedReader(new FileReader(path));

			            String ReadCurrentLineResult;
			            while ((ReadCurrentLineResult = inputStream.readLine()) != null) {
			                //System.out.println(ReadCurrentLineResult);
			                if ((ReadCurrentLineResult
									.contains("Analysis completed at") == true)
									&& AnalysisCompleted == false) {
								// Analysis result ready
								AnalysisCompleted = true;
							}
			            }
			        } finally {
			            if (inputStream != null) {
			                inputStream.close();
			            }
			        }
				}
				robot.delay(2000);
				
				// Enter WPSS results for each Checkpoints
				for (int j = 0; j < 38; j++) {
					BufferedReader inputStream = null;
			        try {
			            inputStream = new BufferedReader(new FileReader(path));			
						// Read WPSS results & enter into Spreadsheet
						String ReadCurrentLine;
						String MsgDesc1Line1 = "";
						String MsgDesc1Line2 = "";
						String MsgDesc1Line3 = "";
						String MsgDesc2Line1 = "";
						String MsgDesc2Line2 = "";
						String MsgDesc2Line3 = "";
						String MsgDesc3Line1 = "";
						String MsgDesc3Line2 = "";
						// String MsgDesc3Line3 = "";
						String MsgDesc4Line1 = "";
						String MsgDesc4Line2 = "";
						// String MsgDesc4Line3 = "";
						String MsgDesc5Line1 = "";
						String MsgDesc5Line2 = "";
						// String MsgDesc5Line3 = "";
						String MsgDesc1 = "";
						String MsgDesc2 = "";
						String MsgDesc3 = "";
						String MsgDesc4 = "";
						String MsgDesc5 = "";
						String FoundIn = "";
						String CommentDesc1 = "";
						String CommentDesc2 = "";
						String CommentDesc3 = "";
						String CommentDesc4 = "";
						String CommentDesc5 = "";
						int linenumber = 0;
						int lineFound = 0;
						int commentLine = 0;
						int NumOfInstance1 = 0;
						int NumOfInstance2 = 0;
						int NumOfInstance3 = 0;
						int NumOfInstance4 = 0;
						int NumOfInstance5 = 0;
						boolean WCAG2_Error1 = false;
						boolean WCAG2_Error2 = false;
						boolean WCAG2_Error3 = false;
						boolean WCAG2_Error4 = false;
						boolean WCAG2_Error5 = false;
						boolean endOfResult = false;
				        while ((ReadCurrentLine = inputStream.readLine()) != null) {
				                //System.out.println(ReadCurrentLine);
							linenumber = linenumber + 1;											
							if ((ReadCurrentLine.contains("Version:") == true)
									&& (ReadCurrentLine.contains("wpss_tool") == true)
									&& (ReadCurrentLine.contains(".pl") == true)) {	
								WPSS_Version = ReadCurrentLine;

								// Extract the WPSS_Tool version #
								WPSS_Version = WPSS_Version.replaceAll(
										"Version:", "");
								WPSS_Version = WPSS_Version.replaceAll(".pl",
										"");
								WPSS_Version = WPSS_Version.replaceAll(
										"wpss_tool", "");
								WPSS_Version = WPSS_Version.replaceAll("_cgi",
										"");
								WPSS_Version = WPSS_Version.replaceAll("_cli",
										"");
								WPSS_Version = WPSS_Version.replaceAll("_en",
										"");
								WPSS_Version = WPSS_Version.replaceAll("_fr",
										"");
								WPSS_Version = WPSS_Version.replaceAll(" ", "");
							}
							if (ReadCurrentLine
									.contains("Results summary table") == true) {
								endOfResult = true;
							}
							if ((ReadCurrentLine.contains(WCAG2StringArray[j]) == true)
									&& (ReadCurrentLine.contains("Testcase") == true)
									// Do not report 4.1.1 G134
									&& (ReadCurrentLine.contains("4.1.1 G134") == false)
									// // Test particular Success Criterion
									// && (WCAG2StringArray[j].contains("3.2.4")
									// == true)
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"  Testcase: ", "FAILED:");
								if (MsgDesc1 == "" == true
										&& endOfResult == false) {
									MsgDesc1 = ReadCurrentLine.toString().trim();
								} else {
									if ((MsgDesc2 == "")
											&& (MsgDesc1
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)) {
										MsgDesc2 = ReadCurrentLine.toString().trim();
									} else {
										if ((MsgDesc3 == "")
												&& (MsgDesc1
														.toString()
														.contentEquals(
																ReadCurrentLine
																		.trim()
																		.toString()) == false)
												&& (MsgDesc2
														.toString()
														.contentEquals(
																ReadCurrentLine
																		.trim()
																		.toString()) == false)) {
											MsgDesc3 = ReadCurrentLine.toString().trim();
										} else {
											if ((MsgDesc4 == "")
													&& ((MsgDesc1
															.toString()
															.contentEquals(
																	ReadCurrentLine
																			.trim()
																			.toString()) == false))
													&& (MsgDesc2
															.toString()
															.contentEquals(
																	ReadCurrentLine
																			.trim()
																			.toString()) == false)
													&& (MsgDesc3
															.toString()
															.contentEquals(
																	ReadCurrentLine
																			.trim()
																			.toString()) == false)) {
												MsgDesc4 = ReadCurrentLine
														.toString().trim();
											} else {
												if ((MsgDesc5 == "")
														&& ((MsgDesc1
																.toString()
																.contentEquals(
																		ReadCurrentLine
																				.trim()
																				.toString()) == false))
														&& (MsgDesc2
																.toString()
																.contentEquals(
																		ReadCurrentLine
																				.trim()
																				.toString()) == false)
														&& (MsgDesc3
																.toString()
																.contentEquals(
																		ReadCurrentLine
																				.trim()
																				.toString()) == false)
														&& (MsgDesc4
																.toString()
																.contentEquals(
																		ReadCurrentLine
																				.trim()
																				.toString()) == false)) {
													MsgDesc5 = ReadCurrentLine
															.toString().trim();
												}
											}
										}
									}
								}
							}

							// First Error message found
							if ((endOfResult == false)
									&& (MsgDesc1.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error1 == false)
									&& (MsgDesc1.toString().contentEquals("") == false)) {
								WCAG2_Error1 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error1 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc1.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc1 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number and
											// column
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											FoundIn = ReadCurrentLine.toString();
											MsgDesc1Line1 = MsgDesc1Line1.toString()
													+ ReadCurrentLine.toString();
											NumOfInstance1 = NumOfInstance1 + 1;
										} else {
											MsgDesc1Line2 = "    "
													+ ReadCurrentLine.toString().trim();
											// System.out.println(MsgDesc1Line2);
											NumOfInstance1 = NumOfInstance1 + 1;
										}
									}
									// Error Message 1 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc1Line2 = "    "
												+ ReadCurrentLine.toString().trim();
									}
									// Error Message 1 Line 3 found
									if ((linenumber == lineFound + 3)
											&& (NumOfInstance1 < 20)) {
										ReadCurrentLine = ReadCurrentLine.toString().trim().replaceAll(
														"\\(line:column\\) ",
														"");
										ReadCurrentLine = ReadCurrentLine.trim().toString().replaceAll("Found \"",
														"\"");
										MsgDesc1Line3 = MsgDesc1Line3.toString() + "In "
												+ FoundIn + " "
												+ ReadCurrentLine.toString().trim() + "\n";
										
									}
									// Error Message 1 Line 3 found
									if ((linenumber == lineFound + 3)
											&& (NumOfInstance1 == 21)) {
										MsgDesc1Line3 = MsgDesc1Line3.toString() + " etc... ";
									}
								} else {
									// Flag it If end of Error message 1
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc1
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc1 == "" == false)) {
										WCAG2_Error1 = false;
									}
								}
							}
							// Error message 2 found
							if ((endOfResult == false)
									&& (MsgDesc2.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error2 == false)
									&& (MsgDesc2.toString().contentEquals("") == false)) {
								WCAG2_Error2 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 2
								if ((linenumber > lineFound)
										&& (WCAG2_Error2 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc2.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc2 == "" == false)) {
									// Error Message 2 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line # and column #
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											MsgDesc2Line1 = MsgDesc2Line1
													+ ReadCurrentLine;
											NumOfInstance2 = NumOfInstance2 + 1;
										} else {
											MsgDesc2Line2 = "    "
													+ ReadCurrentLine.toString().trim();
											NumOfInstance2 = NumOfInstance2 + 1;
										}
									}
									// Error Message 2 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc2Line2 = "    "
												+ ReadCurrentLine.trim();
									}
									// Error Message 2 Line 3 found
									if ((linenumber == lineFound + 3)
											&& (NumOfInstance2 < 20)) {
										ReadCurrentLine = ReadCurrentLine
												.trim().replaceAll(
														"\\(line:column\\) ",
														"");
										ReadCurrentLine = ReadCurrentLine
												.trim().replaceAll("Found \"",
														"\"");
										MsgDesc2Line3 = MsgDesc2Line3 + "In "
												+ FoundIn + " "
												+ ReadCurrentLine.toString().trim() + "\n";
									}
									// Error Message 2 Line 3 found
									if ((linenumber == lineFound + 3)
											&& (NumOfInstance2 == 21)) {
										MsgDesc2Line3 = MsgDesc2Line3 + " etc... ";
									}
								} else {
									// Flag it If end of Error message 2
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc2
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc2 == "" == false)) {
										WCAG2_Error2 = false;
									}
								}
							}

							// Error message 3 found
							if ((endOfResult == false)
									&& (MsgDesc3.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error3 == false)
									&& (MsgDesc3.toString().contentEquals("") == false)) {
								WCAG2_Error3 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 3
								if ((linenumber > lineFound)
										&& (WCAG2_Error3 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc3.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc3 == "" == false)) {
									// Error Message 3 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the Line and Col
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											MsgDesc3Line1 = MsgDesc3Line1
													+ ReadCurrentLine;
											NumOfInstance3 = NumOfInstance3 + 1;
										} else {
											MsgDesc3Line2 = "    "
													+ ReadCurrentLine.trim();
											NumOfInstance3 = NumOfInstance3 + 1;
										}
									}
									// Error Message 3 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc3Line2 = "    "
												+ ReadCurrentLine.toString().trim();
									}
									// Error Message 3 Line 3 found
									if (linenumber == lineFound + 3) {
										// MsgDesc3Line3 =
										// ReadCurrentLine.trim();
									}
								} else {
									// Flag it If end of Error message 1
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc3
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc3 == "" == false)) {
										WCAG2_Error3 = false;
									}
								}
							}

							// Error message 4 found
							if ((endOfResult == false)
									&& (MsgDesc4.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error4 == false)
									&& (MsgDesc4.toString().contentEquals("") == false)) {
								WCAG2_Error4 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error4 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc4.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc4 == "" == false)) {
									// Error Message 4 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number and
											// column
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											MsgDesc4Line1 = MsgDesc4Line1
													+ ReadCurrentLine;
											NumOfInstance4 = NumOfInstance4 + 1;
										} else {
											MsgDesc4Line2 = "    "
													+ ReadCurrentLine.trim();
											NumOfInstance4 = NumOfInstance4 + 1;
										}
									}
									// Error Message 4 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc4Line2 = "    "
												+ ReadCurrentLine.trim();
									}
									// Error Message 4 Line 3 found
									if (linenumber == lineFound + 3) {
										// MsgDesc4Line3 =
										// ReadCurrentLine.trim();
									}
								} else {
									// Flag it If end of Error message 4
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc4
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc4 == "" == false)) {
										WCAG2_Error4 = false;
									}
								}
							}

							// Error message 5 found
							if ((endOfResult == false)
									&& (MsgDesc5.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error5 == false)
									&& (MsgDesc5.toString().contentEquals("") == false)) {
								WCAG2_Error5 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 5
								if ((linenumber > lineFound)
										&& (WCAG2_Error5 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc5.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc5 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number and
											// column
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											MsgDesc5Line1 = MsgDesc5Line1
													+ ReadCurrentLine;
											NumOfInstance5 = NumOfInstance5 + 1;
										} else {
											MsgDesc5Line2 = "    "
													+ ReadCurrentLine.trim();
											NumOfInstance5 = NumOfInstance5 + 1;
										}
									}
									// Error Message 5 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc5Line2 = "    "
												+ ReadCurrentLine.trim();
									}
									// Error Message 5 Line 3 found
									if (linenumber == lineFound + 3) {
										// MsgDesc5Line3 =
										// ReadCurrentLine.trim();
									}
								} else {
									// Flag it If end of Error message 5
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc5
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc5 == "" == false)) {
										WCAG2_Error5 = false;
									}
								}
							}
						}

						if (MsgDesc1 == "" == false) {
							if (MsgDesc1.contains("3.2.4 G197")) {
								if (MsgDesc1Line3 == "" == false) {		
									// Convert meta charset ISO-8859-1 to UTF-8
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#233;", "é");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#192;", "À");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#224;", "à");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#39;", "'");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#232;", "è");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#171;", "«");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#187;" ,"»");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&#201;", "É");
									//MsgDesc1Line3 = MsgDesc1Line3.replaceAll("&amp;", "&");	
									CommentDesc1 = MsgDesc1.toString() + "\n" + MsgDesc1Line2.toString()
											+ "\n" + MsgDesc1Line3 + "\n"
											+ "    Number of Instance: "
											+ NumOfInstance1 + " found in WPSS "
											+ WPSS_Version + "\n\n";
									commentLine = 25;
									//System.out.println("MsgDesc1Line3="+MsgDesc1Line3);
								}
								else {
									CommentDesc1 = MsgDesc1 + "\n" + MsgDesc1Line2
											+ "    Number of Instance: "
											+ NumOfInstance1 + " found in WPSS "
											+ WPSS_Version + "\n\n";
									commentLine = 12;
									//System.out.println(MsgDesc1Line3);	
								}
							} else {
								CommentDesc1 = MsgDesc1 + "\n" + MsgDesc1Line2
										+ "\n"
										+ "    Found in Source Line:Column "
										+ MsgDesc1Line1 + "\n"
										+ "    Number of Instance: "
										+ NumOfInstance1 + " found in WPSS "
										+ WPSS_Version + "\n\n";
								commentLine = 10;
							}
						}
						if (MsgDesc2 == "" == false) {
							CommentDesc2 = MsgDesc2 + "\n" + MsgDesc2Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc2Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance2 + " found in WPSS "
									+ WPSS_Version + "\n\n";
							commentLine = 14;
						}
						if (MsgDesc3 == "" == false) {
							CommentDesc3 = MsgDesc3 + "\n" + MsgDesc3Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc3Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance3 + " found in WPSS "
									+ WPSS_Version + "\n\n";
							commentLine = 21;
						}
						if (MsgDesc4 == "" == false) {
							CommentDesc4 = MsgDesc4 + "\n" + MsgDesc4Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc4Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance4 + " found in WPSS "
									+ WPSS_Version + "\n\n";
							commentLine = 27;
						}
						if (MsgDesc5 == "" == false) {
							CommentDesc5 = MsgDesc5 + "\n" + MsgDesc5Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc5Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance5 + " found in WPSS "
									+ WPSS_Version + "\n";
							commentLine = 36;
						}
						WritableCellFeatures cellFeatures = new WritableCellFeatures();
						cellFeatures.setComment(CommentDesc1 + CommentDesc2
								+ CommentDesc3 + CommentDesc4 + CommentDesc5,
								7, commentLine);
						Label label = new Label(Url_ToBeTest, WCAG2RowArray[j],
								"Fail", arial9formatNoBold);
						label.setCellFeatures(cellFeatures);
						sheet2.addCell(label);
						WCAG2_Error1 = true;

						if ((WCAG2StringArray[j] == "2.4.2")
								&& (WCAG2_Error2_4_2 == true)
								&& (MsgDesc1 == "" == true)) {	
							WritableCellFeatures cellFeatures2_4_2 = new WritableCellFeatures();
							String CurrentComment = ("FAILED 2.4.2 F25: Failure of Success Criterion 2.4.2 due to the title of a Web page not identifying the contents\n"
									+ "For consistency at Environment Canada; Please refer to this EXAMPLE:\n"
									+ "e.g.  www.ec.gc.ca/Air/default.asp?lang=En&n=04104DB7-1\n"
									+ "<title>Environment Canada - Air - Air Quality</title>   PASSED\n"
									+ "Title on this page is >" + TitleStripped + "< FAILED");
							cellFeatures2_4_2.setComment(CurrentComment, 6, 8);
							Label label2_4_2 = new Label(Url_ToBeTest, 34,
									"Fail", arial9formatNoBold);
							label2_4_2.setCellFeatures(cellFeatures2_4_2);
							sheet2.addCell(label2_4_2);
							WCAG2_Error1 = true;
							MsgDesc1 = CurrentComment;
							// System.out.println("2.4.2 Error found ");
						}
						if ((WCAG2StringArray[j] == "2.4.2")
								&& (WCAG2_Error2_4_2 == true)
								&& (MsgDesc1 == "" == false)) {
							WritableCellFeatures cellFeatures2_4_2 = new WritableCellFeatures();
							String CurrentComment = ("FAILED 2.4.2 F25: Failure of Success Criterion 2.4.2 due to the title of a Web page not identifying the contents\n"
									+ "For consistency at Environment Canada; Please refer to this EXAMPLE:\n"
									+ "http://www.ec.gc.ca/Air/default.asp?lang=En&n=04104DB7-1\n"
									+ "<title>Environment Canada - Air - Air Quality</title>   PASSES\n\n"
									+ "Title on this page is >" + TitleStripped + "< FAILED");
							cellFeatures2_4_2.setComment(CommentDesc1
									+ CommentDesc2 + CommentDesc3
									+ CommentDesc4 + CommentDesc5
									+ CurrentComment, 6, 8);
							Label label2_4_2 = new Label(Url_ToBeTest, 34,
									"Fail", arial9formatNoBold);
							label2_4_2.setCellFeatures(cellFeatures2_4_2);
							sheet2.addCell(label2_4_2);
							WCAG2_Error1 = true;
							// System.out.println("2.4.2 Errors found ");
						}

						// The remaining results will pass or N/A
						if (MsgDesc1 == "" == true) {
							if (WCAG2StringArray[j] == "1.2.1"
									|| WCAG2StringArray[j] == "1.2.2"
									|| WCAG2StringArray[j] == "1.2.3"
									|| WCAG2StringArray[j] == "1.2.4"
									|| WCAG2StringArray[j] == "1.2.5"
									|| WCAG2StringArray[j] == "2.2.2"
									|| WCAG2StringArray[j] == "2.3.1"
									|| WCAG2StringArray[j] == "3.3.4") {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										WCAG2RowArray[j], "N/A",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							} else {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										WCAG2RowArray[j], "Pass",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							}
						}
					} // End of all URL While procedures
					catch (Exception ex) {
						JOptionPane pane1 = new JOptionPane(
								"Problem occurs writing WPSS result in Spreadsheet\n\n"
										+ "Error/Warning occurs in line 4370\n\n"
										+ "Please make sure that the script was not interrupted during the process.\n"
										+ ex + "\n\n" + "Press Ok to continue");
						JDialog d1 = pane1.createDialog((JFrame) null,
								"Probem writing result in Spreadsheet file");
						d1.setLocation(400, 500);
						d1.setVisible(true);
					}

					try {
						// Read and write Broken Link results in spreadsheet
						FileInputStream inputFile = new FileInputStream(path2);
						InputStreamReader fr = new InputStreamReader(inputFile,
								"UTF-8");
						// FileReader fr = new FileReader(path2);
						BufferedReader br = new BufferedReader(fr);
						String ReadCurrentLine;
						String MsgDesc1Line1 = "";
						String MsgDesc1Line2 = "";
						String MsgDesc1Line1href = "";
						String MsgDesc1 = "";
						String CommentDesc1 = "";
						int linenumber = 0;
						int lineFound = 0;
						int commentLine = 0;
						int NumOfInstance1 = 0;
						boolean WCAG2_Error1 = false;
						boolean endOfResult = false;
						while ((ReadCurrentLine = br.readLine()) != null) {
							linenumber = linenumber + 1;
							if (ReadCurrentLine
									.contains("Results summary table") == true) {
								endOfResult = true;
							}
							if ((ReadCurrentLine.contains("Testcase") == true)
									&& (ReadCurrentLine.contains("Broken link") == true)
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"  Testcase: ", "FAILED:");
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"Broken link",
										"2.4.5 - G125: Providing links to navigate to related Web pages"
												+ "\nBroken link(s) found");

								if (MsgDesc1 == "" == true
										&& endOfResult == false) {
									MsgDesc1 = ReadCurrentLine.trim();
								}
							}

							if ((ReadCurrentLine.contains("Testcase") == true)
									&& (ReadCurrentLine.contains("broken link") == true)
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"  Testcase: ", "FAILED:");
								ReadCurrentLine = ReadCurrentLine
										.replaceAll(
												"Soft 404 broken link",
												"2.4.5 - G125: Providing links to navigate to related Web pages"
														+ "\nSoft 404 Broken link(s) found");

								if (MsgDesc1 == "" == true
										&& endOfResult == false) {
									MsgDesc1 = ReadCurrentLine.trim();
								}
							}

							// First Error message found
							if ((endOfResult == false)
									&& (MsgDesc1.toString().contentEquals(
											ReadCurrentLine.toString().trim()) == true)
									&& (WCAG2_Error1 == false)
									&& (MsgDesc1.toString().contentEquals("") == false)) {
								WCAG2_Error1 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error1 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.toString().trim()
												.contentEquals("") == false)
										&& (MsgDesc1.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc1 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number and
											// column
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   Line: ", "");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("Column: ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(" :", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":  ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(": ", ":");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll("   ", " ");
											ReadCurrentLine = ReadCurrentLine
													.replaceAll(":-", ":");
											MsgDesc1Line1 = MsgDesc1Line1
													+ ReadCurrentLine;
											MsgDesc1Line1href = ReadCurrentLine;
											NumOfInstance1 = NumOfInstance1 + 1;
										}
									}
									// Error Message 1 Line 2 found
									if (linenumber == lineFound + 2) {
										ReadCurrentLine = ReadCurrentLine
												.replaceAll("href= ", "");
										MsgDesc1Line2 = MsgDesc1Line2
												+ "In line "
												+ MsgDesc1Line1href + " on "
												+ ReadCurrentLine.trim() + "\n";
									}
								} else {
									// Flag it If end of Error message 1
									if ((linenumber > lineFound)
											&& (endOfResult == false)
											&& (ReadCurrentLine.trim()
													.toString()
													.contentEquals("") == true)
											&& (MsgDesc1
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)
											&& (MsgDesc1 == "" == false)) {
										WCAG2_Error1 = false;
									}
								}
							}
						}
						// If broken link were found, the results Failed
						if (MsgDesc1 == "" == false) {
							CommentDesc1 = MsgDesc1 + "\n" + MsgDesc1Line2
									+ "    Number of Instance: "
									+ NumOfInstance1 + " found in WPSS "
									+ WPSS_Version + "\n";
							commentLine = NumOfInstance1 + 4;
							WritableCellFeatures cellFeatures = new WritableCellFeatures();
							cellFeatures.setComment(CommentDesc1, 7,
									commentLine);
							Label label = new Label(Url_ToBeTest, 37, "Fail",
									arial9formatNoBold);
							label.setCellFeatures(cellFeatures);
							sheet2.addCell(label);
							WCAG2_Error1 = true;
						}
						// If no broken link were found, the 2.4.5 results
						// Passed
						if (MsgDesc1 == "" == true) {
							// Result of Success Criterion WCAG2
							WritableCellFeatures cellFeatures = new WritableCellFeatures();
							// cellFeatures.setComment("Failed ",4,2);
							Label label = new Label(Url_ToBeTest, 37, "Pass",
									arial9formatNoBold);
							label.setCellFeatures(cellFeatures);
							sheet2.addCell(label);
						}
						// End of Read Result file from WPSS tool
					} catch (Exception ex) {
						JOptionPane pane1 = new JOptionPane(
								"Problem occurs writing WPSS result in Spreadsheet\n\n"
										+ "Error/Warning occurs in line 4543\n\n"
										+ "Please make sure that the script was not interrupted during the process.\n"
										+ ex + "\n\n" + "Press Ok to continue");
						JDialog d1 = pane1.createDialog((JFrame) null,
								"Probem writing result in Spreadsheet file");
						d1.setLocation(400, 500);
						d1.setVisible(true);

					}
					// End of all URL While procedures

					// Read and write Interoperability results in spreadsheet
					try {
						FileInputStream inputFile = new FileInputStream(path4);
						InputStreamReader fr = new InputStreamReader(inputFile,
								"UTF-8");
						// FileReader fr = new FileReader(path4);
						BufferedReader br = new BufferedReader(fr);
						String ReadCurrentLine;
						String MsgDescSWI_C = "";
						String MsgDescSWI_D = "";
						int linenumber = 0;
						int lineFound = 0;
						int countlineSWI_C = 0;
						int countlineSWI_D = 0;
						// int NumOfInstance1 = 0;
						boolean SWI_C_Error = false;
						boolean SWI_D_Error = false;
						boolean endOfResult = false;
						boolean endOfSWI_C = false;
						boolean endOfSWI_D = false;
						while ((ReadCurrentLine = br.readLine()) != null) {
							linenumber = linenumber + 1;
							if (ReadCurrentLine
									.contains("Results summary table") == true) {
								endOfResult = true;
							}
							if ((ReadCurrentLine.contains("Testcase:") == true)
									&& (ReadCurrentLine.contains("SWI_C") == true)
									&& endOfSWI_C == false
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										" Testcase: ", "FAILED");
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"SWI_C", "");
								MsgDescSWI_C = MsgDescSWI_C
										+ ReadCurrentLine.trim()
										+ " found in WPSS " + WPSS_Version
										+ "\n";
								SWI_C_Error = true;
								lineFound = linenumber;
								endOfSWI_C = false;
							}
							if ((ReadCurrentLine.contains("Testcase:") == true)
									&& (ReadCurrentLine.contains("SWI_D") == true)
									&& endOfSWI_D == false
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										" Testcase: ", "FAILED");
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"SWI_D", "");
								MsgDescSWI_D = MsgDescSWI_D
										+ ReadCurrentLine.trim()
										+ " found in WPSS " + WPSS_Version
										+ "\n";
								SWI_D_Error = true;
								lineFound = linenumber;
								endOfSWI_D = false;
							}
							// SWI_C Error message found
							if ((endOfResult == false)
									&& linenumber > lineFound
									&& endOfSWI_C == false
									&& (SWI_C_Error == true)) {
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										" 	", "");
								MsgDescSWI_C = MsgDescSWI_C
										+ ReadCurrentLine.trim() + "\n";
								countlineSWI_C++;
								if (ReadCurrentLine.length() <= 3) {
									SWI_C_Error = false;
									endOfSWI_C = true;
								}

							}
							// SWI_D Error message found
							if ((endOfResult == false)
									&& linenumber > lineFound
									&& endOfSWI_D == false
									&& (SWI_D_Error == true)) {
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										" 	", "");
								MsgDescSWI_D = MsgDescSWI_D
										+ ReadCurrentLine.trim() + "\n";
								if (ReadCurrentLine.length() <= 3) {
									SWI_D_Error = false;
									countlineSWI_D++;
									endOfSWI_D = true;
								}
							}
						}
						// write result in spreadsheet
						for (int k = 0; k < 35; k++) {
							// If the results Failed
							if ((MsgDescSWI_C == "" == false) && (k == 28)) {
								countlineSWI_C = countlineSWI_C + 3;
								WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
								cellFeatures41.setComment(MsgDescSWI_C, 6,
										countlineSWI_C);
								Label label41 = new Label(Url_ToBeTest,
										InteropRowArray[k], "Fail",
										arial9formatNoBold);
								label41.setCellFeatures(cellFeatures41);
								sheet4.addCell(label41);
								// SWI_C_Error = false;
							} else {
								if (k == 28) {
									// Result of Interoperability
									WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
									Label label41 = new Label(Url_ToBeTest,
											InteropRowArray[k], "Pass",
											arial9formatNoBold);
									label41.setCellFeatures(cellFeatures41);
									sheet4.addCell(label41);
								}
							}
							if ((MsgDescSWI_D == "" == false) && (k == 29)) {
								countlineSWI_D = countlineSWI_D + 3;
								WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
								cellFeatures41.setComment(MsgDescSWI_D, 6,
										countlineSWI_D);
								Label label41 = new Label(Url_ToBeTest,
										InteropRowArray[k], "Fail",
										arial9formatNoBold);
								label41.setCellFeatures(cellFeatures41);
								sheet4.addCell(label41);
								// SWI_D_Error = true;
							} else {
								if (k == 29) {
									// Result of Interoperability
									WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
									Label label41 = new Label(Url_ToBeTest,
											InteropRowArray[k], "Pass",
											arial9formatNoBold);
									label41.setCellFeatures(cellFeatures41);
									sheet4.addCell(label41);
								}
							}

							// If the results Passed
							if (k >= 30) {
								// Result of Interoperability
								WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
								Label label41 = new Label(Url_ToBeTest,
										InteropRowArray[k], "Pass",
										arial9formatNoBold);
								label41.setCellFeatures(cellFeatures41);
								sheet4.addCell(label41);
							}

							// If the results N/A
							// if ((MsgDesc1 == "" == true)&&(k<=27)) {
							if (k <= 27) {
								// Result of Interoperability
								WritableCellFeatures cellFeatures41 = new WritableCellFeatures();
								Label label41 = new Label(Url_ToBeTest,
										InteropRowArray[k], "N/A",
										arial9formatNoBold);
								label41.setCellFeatures(cellFeatures41);
								sheet4.addCell(label41);
							}
							// End of Read/write Result file from WPSS tool
						}
					} catch (Exception ex) {
						JOptionPane pane1 = new JOptionPane(
								"Problem occurs writing Interoperability result in Spreadsheet\n\n"
										+ "Error/Warning occurs in line 4721\n\n"
										+ "Please make sure that the script was not interrupted during the process.\n"
										+ ex + "\n\n" + "Press Ok to continue");
						JDialog d1 = pane1.createDialog((JFrame) null,
								"Probem writing result in Spreadsheet file");
						d1.setLocation(400, 500);
						d1.setVisible(true);
						copy.close();
					}
					// End of all URL While procedures
				}
			}
		}

		// All cells modified/added. Now write out the workbook
		copy.write();
		copy.close();
		robot.delay(1000);
		File file111 = new File("C://QA-WS-Tool//filename.txt");
		file111.delete();

		JOptionPane pane = new JOptionPane(
				"\nCongratulations!!!  The QA Review"
						+ " has been successfully completed with " + Tool
						+ "\n\n" + "Please check your Output file in "
						+ "C:\\QA-WS-Tool\\Output-" + Tool + ".xls\n\n");
		JDialog d = pane.createDialog((JFrame) null,
				"QA Review Successfully Completed\n\n");
		d.setLocation(400, 500);
		d.setVisible(true);
		System.out.println("\n" + "QA Review Successfully Completed.");
		System.exit(1);
	}
}
