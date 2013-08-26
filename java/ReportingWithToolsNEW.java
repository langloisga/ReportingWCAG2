package Automation;

import resources.Automation.ReportingWithToolsNEWHelper;
import com.rational.test.ft.*;
import com.rational.test.ft.object.interfaces.*;
import com.rational.test.ft.object.interfaces.SAP.*;
import com.rational.test.ft.object.interfaces.WPF.*;
import com.rational.test.ft.object.interfaces.dojo.*;
import com.rational.test.ft.object.interfaces.siebel.*;
import com.rational.test.ft.object.interfaces.flex.*;
import com.rational.test.ft.object.interfaces.generichtmlsubdomain.*;
import com.rational.test.ft.script.*;
import com.rational.test.ft.value.*;
import com.rational.test.ft.vp.*;
import com.ibm.rational.test.ft.object.interfaces.sapwebportal.*;

// BEGIN custom imports 		
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Orientation;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;

import java.awt.AWTException;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.List;
import java.awt.TextComponent;
import java.awt.TextField;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.ClipboardOwner;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.FlavorListener;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.event.AWTEventListener;
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
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JRadioButton;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;

/**
 * <b>Functional Test Script</b> <b>Description: </b> Will populate any tool
 * results in Spreadsheet Report
 * 
 * 
 * @author Gaston Langlois
 * @since 2012/06/20
 */

public class ReportingWithToolsNEW extends ReportingWithToolsNEWHelper {

	/**
	 * Script Name : <b>ReportingWithTools</b> Generated : <b>2013-01-30 9:38:39
	 * AM</b> Description : Functional Test Script Original Host : WinNT Version
	 * 5.1 Build 2600 (S)
	 * 
	 * @throws WriteException
	 * @throws RowsExceededException
	 * @throws IOException
	 * @throws JXLException
	 * @throws AWTException
	 * @throws UnsupportedFlavorException
	 * @throws HeadlessException
	 */
	public void testMain(Object[] args) throws RowsExceededException,
			WriteException, IOException, JXLException, AWTException,
			HeadlessException, UnsupportedFlavorException {

		File file = new File("m://filename.txt");
		file.delete();
 
		// Close Internet Explore if opened
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
		sleep(1);

		final JFrame f1 = new JFrame("Automation Program Application");
		// f.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		f1.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		f1.setSize(400, 200);
		f1.setLocation(300, 300);

		// f1.addWindowListener(new WindowAdapter() {
		// public void windowClosing(WindowEvent we) {
		// //System.exit(0);
		// //new ClosingFrame();
		// f1.setVisible(false);
		// }
		// });
		Robot robot = new Robot();
		
		JTextField tf;
		final JPanel entreePanel = new JPanel();
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
		
		//Enter Input file here
		
		entreePanel.add(new JLabel("Input File:  "));	
		entreePanel.setFont(new java.awt.Font("Arial", 14, 14));
		entreePanel.add(new JTextField("C:\\Test7.xls",39));
		entreePanel.setFont(new java.awt.Font("Arial", 14, 14));
		
		
		
		
		//testLabel.add(Box.createHorizontalStrut(2));
		//testLabel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

		// final JPanel condimentsPanel = new JPanel();
		//condimentsPanel.add(new JCheckBox("Direct Input"));
		// condimentsPanel.add(new JCheckBox("Live"));
		final JPanel entreePanel2 = new JPanel();
		final ButtonGroup entreeGroup2 = new ButtonGroup();
		// Add a text message to select the tool to run
		String text2 = "Select the type of Source code:      ";
		text2 += "\n";
		
		JLabel testLabel2 = new JLabel(text2);
		// customize radio button input
		entreePanel2.add(testLabel2);
		JRadioButton radioButton2;
		entreePanel2.add(radioButton2 = new JRadioButton("Direct Input"));
		radioButton2.setActionCommand("Direct Input");
		radioButton2.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup2.add(radioButton2);
		entreePanel2.add(radioButton2 = new JRadioButton("Url", true));
		radioButton2.setActionCommand("Url");
		// Preselect the Live radio button
		// radioButton2.setSelected(true);
		radioButton2.setFont(new java.awt.Font("Arial", 0, 14));
		entreeGroup2.add(radioButton2);
		testLabel2.setFont(new java.awt.Font("Arial", 0, 14));
		entreePanel2.add(new JLabel("Output File:"));	
		entreePanel2.add(new JTextField("C:\\Output.xls",39));
		
		
		JPanel orderPanel = new JPanel();
		JButton orderButton = new JButton("Submit");
		orderButton.setFont(new java.awt.Font("Arial", 0, 14));
		orderPanel.add(orderButton);		
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
					writer = new PrintWriter("m://filename.txt", "UTF-8");
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (UnsupportedEncodingException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				//Capture the radio button choices
				writer.println(entree);
				String entree2 = entreeGroup2.getSelection().getActionCommand();
				writer.println(entree2);

				//Return the value of the button name
				//String cmd = ae.getActionCommand();  
				//System.out.println("CMD IN ECHO LISTENER = " + cmd);  
	
				//String TEST = JTextField.getFocusedComponent();
				//String TEST = ae.getSource().toString();
				//System.out.println("CMD IN ECHO LISTENER = " + TEST);
	
				//System.out.println(returnValues);
				//for (int i = 0; i < components.length; i++) {
				//	returnValues[0] = jTextField1.getText();	
				//}

				//for (int i = 0; i < components.length; i++) {
				//	JCheckBox cb = (JCheckBox) components[i];
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

		File f = new File("m://filename.txt");
		while (f.exists() == false) {
			robot.delay(3000);
		}

		
		String Tool = "";
		String SourceCode = "";
		try {
			// Read and copy file
			FileReader frSource = new FileReader("M:\\filename.txt");
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
		}

		// EXIT RFT Script
		// System.exit(0);
		System.out.println("Input  spreadsheet copied  from C:\\Test7.xls");
		Workbook workbook = Workbook.getWorkbook(new File("C:///Test7.xls"));
		WritableWorkbook copy = Workbook.createWorkbook(new File(
				"C:///output.xls"), workbook);
		WritableSheet sheet2 = copy.getSheet(2);

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
		//Runtime.getRuntime().exec("taskkill /F /IM excel.exe ");
		System.out.println("Output spreadsheet generated to C:\\output.xls");

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

		// Declare array for WCAG2 Sufficient Techniques Strings in the
		// spreasheet
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
		// for (int i=1; i<69; i++)
		for (int i = 1; i < 69; i++) {
			Url_ToBeTest = i;
			// Initialize variables for each url and WCAG

			String path = "C:\\Program Files\\WPSS_Tool\\results\\WorkingStorage_acc.txt";
			String path2 = "C:\\Program Files\\WPSS_Tool\\results\\WorkingStorage_link.txt";
			String path3 = "C:\\Documents and Settings\\langloisga\\Application Data\\AI Internet Solutions\\CSE HTML Validator\\12.0\\batchreport1.html";
			String WorkingStorage = "C:\\Program Files\\WPSS_Tool\\results\\WorkingStorage";
			WritableCell cell = sheet2.getWritableCell(Url_ToBeTest, 5);

			boolean AnalysisCompleted = false;
			
			// Get page entry
			String str = cell.getContents();
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

			
			if (Tool.contains("CSE")) {
				// TODO Insert here CSE code to saving the result file
				// Close CSE if already opened
				Runtime.getRuntime().exec("taskkill /F /IM cse120.exe ");
				sleep(2);
				// Open the CSE HTML Validator Pro Batch Wizard
				Runtime.getRuntime()
						.exec("cmd /c start C:\\\"Program Files\\CSE HTMLValidator\\cmdlineprocessor.exe");
				sleep(2);
				// Window: cse120.exe: CSE HTML Validator Pro v12.01 -
				csehtmlValidatorProV1202CProgr().click();
				csehtmlValidatorProV1202CProgr().inputKeys("{F2}");
			

				
				//Copy content in the Clipboard
				Toolkit toolkit = Toolkit.getDefaultToolkit();
				Clipboard clipboard = toolkit.getSystemClipboard();
				StringSelection strSel = new StringSelection(cell.getContents());
				clipboard.setContents(strSel, null);				

				// Enter url from the spreadsheet			
				addmenu3().click(atPoint(14, 11));
				addpopupMenu3().click(atPath("Add URL"));
				enterURLTargetToAddwindow3().inputKeys("{BKSP}^v");
				okbutton10().click();

				//Click on the Ok from CSE  400 horizontal, 480 vertical
				robot.mouseMove(400, 480); 
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(2000);
				
				// For each html click on Startbutton
				startbutton3().click(atPoint(24, 12));
				robot.delay(8000);
				
				// If CSE Results browser still do not exists
				if (browser_htmlBrowser().exists()== false) {
					Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
					startbutton2().click(atPoint(24, 12));
					robot.delay(9000);	
					// Minimize Browser
					browser_htmlBrowser(document_batchWizardReportCSEH(),DEFAULT_FLAGS).minimize();	
				} 
				else {
					// Minimize Browser
					browser_htmlBrowser(document_batchWizardReportCSEH(),DEFAULT_FLAGS).minimize();
				}
						
				// Window: cse120.exe: CSE HTML Validator Pro Batch Wizard
				csehtmlValidatorProBatchWizard3().exists();
				robot.delay(1000);
				// Click on the Target List cursor 90 horizontal and 130
				// vertical
				robot.mouseMove(90, 130);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);

				// Select all
				csehtmlValidatorProBatchWizard3().inputKeys("^a");
				deletebutton6().click(atPoint(25, 9));
				robot.delay(100);
				deletebutton7().click(atPoint(36, 14));
				// deletebutton().click(atPoint(23, 12));
				Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");

				for (int j = 0; j < 38; j++) {
					try {
						// Read WPSS results & enter it to the spreadsheet
						FileReader fr = new FileReader(path3);
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
									&& (ReadCurrentLine
											.contains("Accessibility Error") == true)) {
								String str1 = ReadCurrentLine;
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
								int PosTrimDesc = MsgDesc.indexOf(" [A");
								if ((MsgPos < PosTrimDesc)
										&& (PosTrimDesc < PositionVisit)) {
									MsgDesc = String.copyValueOf(charArray,
											MsgPos + 16, PosTrimDesc - MsgPos);
								} else {
									if (Position < PositionVisit) {
										MsgDesc = String.copyValueOf(charArray,
												MsgPos + 16, Position - MsgPos
														- 16);
									} else {
										MsgDesc = String.copyValueOf(charArray,
												MsgPos + 16, PositionVisit
														- MsgPos - 17);
									}
								}
								ErrorLocation = String.copyValueOf(charArray,
										33, LocPos - 49);
								// System.out.println("SuffTech="+SuffTech+" ErrorLocation="+ErrorLocation);
								if (MsgDesc1 == "" == true) {
									MsgDesc1 = SuffTech.trim();
									MsgDesc1Line2 = MsgDesc.trim();
									// System.out.println("MsgDesc1="+MsgDesc1);
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
										SuffTech.trim().toString()) == true)
										&& (ReadCurrentLine
												.contains("Accessibility Error") == true)
										&& (MsgDesc1.toString().contentEquals(
												"") == false)) {
									MsgDesc1 = SuffTech.trim();
									MsgDesc1Line2 = MsgDesc.trim();
									MsgDesc1Line1 = MsgDesc1Line1
											+ ErrorLocation + "; ";
									NumOfInstance1 = NumOfInstance1 + 1;
								} else {
									if ((MsgDesc2.toString().contentEquals(
											SuffTech.trim().toString()) == true)
											&& (MsgDesc1
													.toString()
													.contentEquals(
															SuffTech.trim()
																	.toString()) == false)) {
										MsgDesc2 = SuffTech.trim();
										MsgDesc2Line2 = MsgDesc.trim();
										MsgDesc2Line1 = MsgDesc2Line1
												+ ErrorLocation + "; ";
										NumOfInstance2 = NumOfInstance2 + 1;
									} else {
										if ((MsgDesc3.toString().contentEquals(
												SuffTech.trim().toString()) == true)
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
								CommentDesc1 = MsgDesc1 + " " + MsgDesc1Line2
										+ "\n"
										+ "    Found in Source Line:Column "
										+ MsgDesc1Line1 + "\n"
										+ "    Number of Instance: "
										+ NumOfInstance1 + "\n\n";
								commentLine = 10;
							}
						}
						if (MsgDesc2 == "" == false) {
							CommentDesc2 = MsgDesc2 + " " + MsgDesc2Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc2Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance2 + "\n\n";
							commentLine = 20;
						}
						if (MsgDesc3 == "" == false) {
							CommentDesc3 = MsgDesc3 + " " + MsgDesc3Line2
									+ "\n" + "    Found in Source Line: "
									+ MsgDesc3Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance3 + "\n\n";
							commentLine = 30;
						}
						if (MsgDesc4 == "" == false) {
							CommentDesc4 = MsgDesc4 + " " + MsgDesc4Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc4Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance4 + "\n\n";
							commentLine = 40;
						}
						if (MsgDesc5 == "" == false) {
							CommentDesc5 = MsgDesc5 + " " + MsgDesc5Line2
									+ "\n" + "    Found in Source Line:Column "
									+ MsgDesc5Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance5;
							commentLine = 44;
						}
						WritableCellFeatures cellFeatures = new WritableCellFeatures();
						cellFeatures.setComment(CommentDesc1 + CommentDesc2
								+ CommentDesc3 + CommentDesc4 + CommentDesc5,
								6, commentLine);
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
					}
				}
				// CSE HTML Validator Pro Batch Wizard
				//startbutton2().click(atPoint(18, 7));

				// Window: IEXPLORE.EXE: Windows Internet Explorer
				//windowsInternetExplorerwindow().minimize();
			}
			
			if ((Tool.contains("WPSS") || Tool.contains("W3C"))
					&& SourceCode.contains("Direct Input")) {
				if (Tool.contains("WPSS")) {
					// Close PWGSC WPSS tool if already opened
					Runtime.getRuntime().exec("taskkill /F /IM perl.exe ");
					sleep(2);
					// Open the PWGSC WPSS tool
					Runtime.getRuntime()
							.exec("cmd /c start C:\\\"Program Files\\WPSS_Tool\\wpss_tool.pl");
				}
				File file2 = new File("m://test.html");
				file2.delete();
				// Go to firefox (Home page Environment Canade)
				if (environnementCanadaEnvironment().exists()) {
					environnementCanadaEnvironment().exists();
					comboListBoxcomboBox3().click(atPoint(420, 19));
					robot.mouseMove(600, 90);
					robot.delay(300);
					robot.keyPress(KeyEvent.VK_DELETE);
					robot.keyRelease(KeyEvent.VK_DELETE);
					robot.delay(100);
					//environnementCanadaEnvironment().inputKeys(str.toString());

					// Copy content in the Clipboard
					Toolkit toolkit = Toolkit.getDefaultToolkit();
					Clipboard clipboard = toolkit.getSystemClipboard();
					StringSelection strSel = new StringSelection(cell.getContents());
					clipboard.setContents(strSel, null);

					// Press Ctrl+V to enter url
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.delay(500);
				}

				// Press Reload current page
				robot.mouseMove(840, 90);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				// Wait 6 seconds assuming that the page has been loaded
				robot.delay(6000);

				// Press View Source from Firefox
				robot.mouseMove(750, 170); // Click on View Source button
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(1000);

				// Click on View Source drop down
				robot.mouseMove(750, 240);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(4000);

				// Press Ctrl+S to save the html file and press save button
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_S);
				robot.delay(1000);

				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_SEMICOLON);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.keyRelease(KeyEvent.VK_SEMICOLON);
				//robot.keyPress(KeyEvent.VK_BACK_SLASH);
				//robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyPress(KeyEvent.VK_E);
				robot.keyRelease(KeyEvent.VK_E);
				robot.keyPress(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_S);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyPress(KeyEvent.VK_PERIOD);
				robot.keyRelease(KeyEvent.VK_PERIOD);
				robot.keyPress(KeyEvent.VK_H);
				robot.keyRelease(KeyEvent.VK_H);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);

				// Save the htm file
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				// Wait 1 second assuming the htm file is downloaded
				robot.delay(1000);

				// Click Yes to overwrite
				robot.mouseMove(590, 570); // Click on Yes button
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				// Wait 1 second assuming the htm file is downloaded
				robot.delay(1000);

				// Close Source window
				robot.mouseMove(1265, 10); // Click on Close window
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				// Wait 3 second assuming the htm file has been downloaded
				robot.delay(3000);

				// Click on the Home to prepare for next url
				robot.mouseMove(1150, 90);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(700);

				// Click on Close Tab in firefox
				robot.mouseMove(235, 60);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.delay(200);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(200);

				String contentSource = "";
				try {
					// Read and copy file
					FileReader frSource = new FileReader("M:\\test.html");
					BufferedReader brSource = new BufferedReader(frSource);
					String ReadCurrentLineSource;
					int linenumberSource = 0;
					// Reading WPSS results
					while ((ReadCurrentLineSource = brSource.readLine()) != null) {
						linenumberSource = linenumberSource + 1;
						contentSource = contentSource
								+ ReadCurrentLineSource + "\n";
					}

					// Copy content in the Clipboard
					Toolkit toolkit = Toolkit.getDefaultToolkit();
					Clipboard clipboard = toolkit.getSystemClipboard();
					StringSelection strSel = new StringSelection(contentSource);
					clipboard.setContents(strSel, null);
					// System.out.println(contentSource);
				} // End of Read and Copy file

				catch (Exception ex) {
				}
			}
			
			if (Tool.contains("WPSS")
					&& SourceCode.contains("Direct Input")) {
				// Direct HTML Input
				if (Tool.contains("WPSS")) {
					// Close PWGSC WPSS tool if already opened
					Runtime.getRuntime().exec("taskkill /F /IM perl.exe ");
					robot.delay(2000);
					// Open the PWGSC WPSS tool
					Runtime.getRuntime()
							.exec("cmd /c start C:\\\"Program Files\\WPSS_Tool\\wpss_tool.pl");
					robot.delay(2000);
				}
				pagetablistpageTabList3().click(
						atName("Direct HTML Input"), atPoint(33, 9));
				edittext6().click(atPoint(219, 47));
				pwgscwpssValidatorwindow().inputKeys("^v");
				checkURLListbutton().click(atPoint(87, 11));
			}
			
			if (Tool.contains("W3C")) {
				// Open W3C browser at http://html5.validator.nu
				if (SourceCode.contains("Direct Input")) {
					startApp("W3C_Markup_Direct");
					robot.delay(5000);
					browser_htmlBrowser(document_theW3CMarkupValidatio2(),
							DEFAULT_FLAGS).maximize();

					// Click on the text box to paste the html code
					robot.mouseMove(650, 275);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(500);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(3000);

					// Press Ctrl+S to save the html file
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);
					robot.delay(1000);

					// click on the Check button
					robot.mouseMove(650, 430);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(9000);
				}

				// Open W3C browser at http://html5.validator.nu
				if (SourceCode.contains("Url")) {
					startApp("W3C_Markup_Url");
					robot.delay(5000);
					browser_htmlBrowser(document_theW3CMarkupValidatio2(),
							DEFAULT_FLAGS).maximize();

					// Copy content in the Clipboard
					Toolkit toolkit = Toolkit.getDefaultToolkit();
					Clipboard clipboard = toolkit.getSystemClipboard();
					StringSelection strSel = new StringSelection(cell.getContents());
					clipboard.setContents(strSel, null);

					robot.mouseMove(300, 335);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(500);

					// Press Ctrl+S to save the html file
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyRelease(KeyEvent.VK_V);
					robot.delay(500);

					// click on the Check button
					robot.mouseMove(650, 440);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.mousePress(InputEvent.BUTTON1_MASK);
					robot.mouseRelease(InputEvent.BUTTON1_MASK);
					robot.delay(15000);
				}

				// View Source from iexplore
				robot.mouseMove(1000, 200); // Right-Click on page to view
											// source
				robot.mousePress(InputEvent.BUTTON3_MASK);
				robot.mouseRelease(InputEvent.BUTTON3_MASK);
				robot.delay(1000);

				// Click on View Source drop down
				robot.mouseMove(1030, 400);
				robot.mousePress(InputEvent.BUTTON3_MASK);
				robot.mouseRelease(InputEvent.BUTTON3_MASK);
				robot.delay(2000); // Wait 2 second assuming the source
									// displayed

				// Save source code from notepad
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_F);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_F);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_A);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.keyRelease(KeyEvent.VK_A);
				robot.delay(500);

				robot.keyPress(KeyEvent.VK_BACK_SPACE);
				robot.keyRelease(KeyEvent.VK_BACK_SPACE);
				robot.keyPress(KeyEvent.VK_M);
				robot.keyRelease(KeyEvent.VK_M);
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(KeyEvent.VK_SEMICOLON);
				robot.keyRelease(KeyEvent.VK_SHIFT);
				robot.keyRelease(KeyEvent.VK_SEMICOLON);
				robot.keyPress(KeyEvent.VK_BACK_SLASH);
				robot.keyRelease(KeyEvent.VK_BACK_SLASH);
				robot.keyPress(KeyEvent.VK_W);
				robot.keyRelease(KeyEvent.VK_W);
				robot.keyPress(KeyEvent.VK_3);
				robot.keyRelease(KeyEvent.VK_3);
				robot.keyPress(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_C);
				robot.keyPress(KeyEvent.VK_PERIOD);
				robot.keyRelease(KeyEvent.VK_PERIOD);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);
				robot.keyPress(KeyEvent.VK_X);
				robot.keyRelease(KeyEvent.VK_X);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_T);

				// Press the Save button from notepad
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_S);
				robot.delay(1000);

				// Overwrite file
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_Y);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_Y);
				robot.delay(500);

				// Exit Notepad
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_F);
				robot.keyRelease(KeyEvent.VK_F);
				robot.keyPress(KeyEvent.VK_X);
				robot.keyRelease(KeyEvent.VK_X);
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.delay(500);

				// Close the Internet Explore browser
				Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe ");
				robot.delay(1000);

				// Click on the Home to prepare for next url
				robot.mouseMove(1155, 90);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(2000);

				// Click on Close Tab in firefox
				// robot.mouseMove(235, 60); seems to be the same as the
				// position in EditorPagetitle
				robot.mouseMove(242, 60);
				robot.mousePress(InputEvent.BUTTON1_MASK);
				robot.mouseRelease(InputEvent.BUTTON1_MASK);
				robot.delay(500);

				try {
					// Read 4.1.1 results & enter them in spreadsheet
					FileReader fr4_1_1 = new FileReader("m:\\w3c.txt");
					BufferedReader br4_1_1 = new BufferedReader(fr4_1_1);
					String content4_1_1 = "";
					String ReadCurrentLine4_1_1;
					int linenumber4_1_1 = 0;
					int ErrorFound = 0;
					int WCAG2_4_1_1Row = 54;
					boolean WCAG2_4_1_1Error1 = false;
					String CurrentComment = "";
					while ((ReadCurrentLine4_1_1 = br4_1_1.readLine()) != null)
					// WCAG2.0 4.1.1
					{
						linenumber4_1_1 = linenumber4_1_1 + 1;
						content4_1_1 = content4_1_1 + ReadCurrentLine4_1_1
								+ "\n";
						if ((ReadCurrentLine4_1_1.contains("   <td colspan=") == true)
								&& ReadCurrentLine4_1_1.contains("invalid")
								&& WCAG2_4_1_1Error1 == false) {
							// Failure of Success Criterion WCAG2 4.1.1
							WCAG2_4_1_1Error1 = true;
							ErrorFound = linenumber4_1_1 + 1;
						}
						if ((WCAG2_4_1_1Error1 == true)
								&& linenumber4_1_1 == ErrorFound) {
							// enter error
							WritableCellFeatures cellFeatures4_1_1 = new WritableCellFeatures();
							CurrentComment = ("FAILED 4.1.1 Parsing G134: Validating Web the page using the W3C Markup validaton Service at http://validator.w3.org/#validate-by-input and/or http://validator.w3.org/#validate-by-uri\n" + ReadCurrentLine4_1_1);
							// System.out.println("Comment="+cellFeatures4_1_1.getComment());
							cellFeatures4_1_1.setComment(CurrentComment, 5, 6);
							Label label4_1_1 = new Label(Url_ToBeTest,
									WCAG2_4_1_1Row, "Fail", arial9formatNoBold);
							label4_1_1.setCellFeatures(cellFeatures4_1_1);
							// System.out.println("Comment="+cellFeatures4_1_1.getComment());
							sheet2.addCell(label4_1_1);
						}
					} // end of while procedure for 4.1.1

					// System.out.println("WCAG2_4_1_1Error1="+WCAG2_4_1_1Error1);
					// The remaining results will pass or N/A
					for (int j = 0; j < 38; j++) {
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
									WCAG2RowArray[j], "N/A", arial9formatNoBold);
							label2.setCellFeatures(cellFeatures2);
							sheet2.addCell(label2);
						} else {
							if (WCAG2StringArray[j] != "4.1.1") {
								// Result of Success Criterion WCAG2
								WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
								// cellFeatures.setComment("Failed ",4,2);
								Label label2 = new Label(Url_ToBeTest,
										WCAG2RowArray[j], "Pass",
										arial9formatNoBold);
								label2.setCellFeatures(cellFeatures2);
								sheet2.addCell(label2);
							}
							if ((WCAG2StringArray[j] == "4.1.1")
									&& (WCAG2_4_1_1Error1 == false)) {
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
					}

				} // End of all URL While procedures
				catch (Exception ex) {
				}

			}

			if (Tool.contains("WPSS") && SourceCode.contains("Url")) {
				if (Tool.contains("WPSS")) {
					// Close PWGSC WPSS tool if already opened
					Runtime.getRuntime().exec("taskkill /F /IM perl.exe ");
					robot.delay(2000);
					// Open the PWGSC WPSS tool
					Runtime.getRuntime()
							.exec("cmd /c start C:\\\"Program Files\\WPSS_Tool\\wpss_tool.pl");
					robot.delay(3000);
				}
				// Window: perl.exe: PWGSC WPSS Validator
				pagetablistpageTabList3().click(atName("URL List"),
						atPoint(28, 8));
				robot.delay(3000);
				edittext6().click(atPoint(29, 28));
				pwgscwpssValidatorwindow().inputKeys(cell.getContents());
				checkURLListbutton().click(atPoint(77, 8));
			}
			AnalysisCompleted = false;
			// Save the result every 3 seconds until the analysis is completed
			if (Tool.contains("WPSS")) {
				while (AnalysisCompleted == false) {
					// Goto the Results Window
					// Window: perl.exe: Results Window
					pagetablistpageTabList().click(atName("ACC"),
							atPoint(18, 9));
					resultsWindowwindow().click(atPoint(16, 38));
					resultsWindowpopupMenu().click(atPath("Save As"));
					comboBoxcomboBox().click(atPoint(43, 9));
					saveAswindow().inputChars(WorkingStorage);
					savebutton().click(atPoint(26, 8));
					String contentResult = "";
					
					// Wait 3 seconds assuming that the file was successfully saved
					robot.delay(3000);
					try {
						// Read Source file
						FileReader frResult = new FileReader(path);
						BufferedReader brResult = new BufferedReader(frResult);
						String ReadCurrentLineResult;
						int linenumberSource = 0;
						// Reading WPSS results
						while ((ReadCurrentLineResult = brResult.readLine()) != null) {
							linenumberSource = linenumberSource + 1;
							contentResult = contentResult
									+ ReadCurrentLineResult + "\n";
							if ((ReadCurrentLineResult
									.contains("Analysis completed at") == true)
									&& AnalysisCompleted == false) {
								// Analysis result ready
								AnalysisCompleted = true;
							}
						}
						// System.out.println(contentResult);
						// Window: perl.exe: PWGSC WPSS Validator
						pwgscwpssValidatorwindow().move(atPoint(1, 104));
					}
					// End of Checking Results file to see if Analysis is
					// completed/
					catch (Exception ex) {
					}
				}

				// Enter WPSS results for each Checkpoints
				for (int j = 0; j < 38; j++) {

					try {
						// Read WPSS results & enter them into Result
						// spreadsheet

						FileReader fr = new FileReader(path);
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
						while ((ReadCurrentLine = br.readLine()) != null) {
							linenumber = linenumber + 1;
							if (ReadCurrentLine
									.contains("Results summary table") == true) {
								endOfResult = true;
							}
							if ((ReadCurrentLine.contains(WCAG2StringArray[j]) == true)
									&& (ReadCurrentLine.contains("Testcase") == true)
									// /////////// test
									// && (WCAG2StringArray[j].contains("1.3.1")
									// ==
									// true)
									// ////////// test
									&& endOfResult == false) {
								// Got the line number here
								ReadCurrentLine = ReadCurrentLine.replaceAll(
										"  Testcase: ", "FAILED:");
								if (MsgDesc1 == "" == true
										&& endOfResult == false) {
									MsgDesc1 = ReadCurrentLine.trim();
								} else {
									if ((MsgDesc2 == "")
											&& (MsgDesc1
													.toString()
													.contentEquals(
															ReadCurrentLine
																	.trim()
																	.toString()) == false)) {
										MsgDesc2 = ReadCurrentLine.trim();
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
											MsgDesc3 = ReadCurrentLine.trim();
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
														.trim();
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
															.trim();
												}
											}
										}
									}
								}
							}

							// First Error message found
							if ((endOfResult == false)
									&& (MsgDesc1.toString().contentEquals(
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error1 == false)
									&& (MsgDesc1.toString().contentEquals("") == false)) {
								WCAG2_Error1 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error1 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc1.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc1 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc1Line1 = MsgDesc1Line1
													+ String.copyValueOf(
															charArray, 10,
															Position - 11)
													+ "; ";
											NumOfInstance1 = NumOfInstance1 + 1;
										} else {
											MsgDesc1Line2 = "    "
													+ ReadCurrentLine.trim();
											NumOfInstance1 = NumOfInstance1 + 1;
										}
									}
									// Error Message 1 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc1Line2 = "    "
												+ ReadCurrentLine.trim();
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
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error2 == false)
									&& (MsgDesc2.toString().contentEquals("") == false)) {
								WCAG2_Error2 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 2
								if ((linenumber > lineFound)
										&& (WCAG2_Error2 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc2.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc2 == "" == false)) {
									// Error Message 2 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc2Line1 = MsgDesc2Line1
													+ String.copyValueOf(
															charArray, 10,
															Position - 11)
													+ "; ";
											NumOfInstance2 = NumOfInstance2 + 1;
										} else {
											MsgDesc2Line2 = "    "
													+ ReadCurrentLine.trim();
											NumOfInstance2 = NumOfInstance2 + 1;
										}
									}
									// Error Message 2 Line 2 found
									if (linenumber == lineFound + 2) {
										MsgDesc2Line2 = "    "
												+ ReadCurrentLine.trim();
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
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error3 == false)
									&& (MsgDesc3.toString().contentEquals("") == false)) {
								WCAG2_Error3 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 3
								if ((linenumber > lineFound)
										&& (WCAG2_Error3 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc3.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc3 == "" == false)) {
									// Error Message 3 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc3Line1 = MsgDesc3Line1
													+ String.copyValueOf(
															charArray, 10,
															Position - 11)
													+ "; ";
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
												+ ReadCurrentLine.trim();
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
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error4 == false)
									&& (MsgDesc4.toString().contentEquals("") == false)) {
								WCAG2_Error4 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error4 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc4.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc4 == "" == false)) {
									// Error Message 4 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc4Line1 = MsgDesc4Line1
													+ String.copyValueOf(
															charArray, 10,
															Position - 11)
													+ "; ";
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
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error5 == false)
									&& (MsgDesc5.toString().contentEquals("") == false)) {
								WCAG2_Error5 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 5
								if ((linenumber > lineFound)
										&& (WCAG2_Error5 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc5.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc5 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc5Line1 = MsgDesc5Line1
													+ String.copyValueOf(
															charArray, 10,
															Position - 11)
													+ "; ";
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
							if (MsgDesc1Line2.contains("Fails validation,")) {
								CommentDesc1 = MsgDesc1
										+ "\n"
										+ "    Each web page should have no error from the W3C Markup validaton Service."
										+ "  HTML source should be validated at http://validator.w3.org/#validate-by-input and/or http://validator.w3.org/#validate_by_uri"
										+ "\n\n";
								commentLine = 8;
							} else {
								CommentDesc1 = MsgDesc1 + "\n" + MsgDesc1Line2
										+ "\n" + "    Found in Source Line: "
										+ MsgDesc1Line1 + "\n"
										+ "    Number of Instance: "
										+ NumOfInstance1 + "\n\n";
								commentLine = 8;
							}
						}
						if (MsgDesc2 == "" == false) {
							CommentDesc2 = MsgDesc2 + "\n" + MsgDesc2Line2
									+ "\n" + "    Found in Source Line: "
									+ MsgDesc2Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance2 + "\n\n";
							commentLine = 14;
						}
						if (MsgDesc3 == "" == false) {
							CommentDesc3 = MsgDesc3 + "\n" + MsgDesc3Line2
									+ "\n" + "    Found in Source Line: "
									+ MsgDesc3Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance3 + "\n\n";
							commentLine = 21;
						}
						if (MsgDesc4 == "" == false) {
							CommentDesc4 = MsgDesc4 + "\n" + MsgDesc4Line2
									+ "\n" + "    Found in Source Line: "
									+ MsgDesc4Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance4 + "\n\n";
							commentLine = 27;
						}
						if (MsgDesc5 == "" == false) {
							CommentDesc5 = MsgDesc5 + "\n" + MsgDesc5Line2
									+ "\n" + "    Found in Source Line: "
									+ MsgDesc5Line1 + "\n"
									+ "    Number of Instance: "
									+ NumOfInstance5;
							commentLine = 36;
						}
						WritableCellFeatures cellFeatures = new WritableCellFeatures();
						cellFeatures.setComment(CommentDesc1 + CommentDesc2
								+ CommentDesc3 + CommentDesc4 + CommentDesc5,
								6, commentLine);
						Label label = new Label(Url_ToBeTest, WCAG2RowArray[j],
								"Fail", arial9formatNoBold);
						label.setCellFeatures(cellFeatures);
						sheet2.addCell(label);
						WCAG2_Error1 = true;

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
					}

					try {
						// Read WPSS results & enter them into Result
						// spreadsheet
						FileReader fr = new FileReader(path2);
						BufferedReader br = new BufferedReader(fr);
						String ReadCurrentLine;
						String MsgDesc1Line1 = "";
						String MsgDesc1Line2 = "";
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
										"2.4.5 - G126: Providing a list of links to all other Web pages"
												+ "\nBroken link(s)found");

								if (MsgDesc1 == "" == true
										&& endOfResult == false) {
									MsgDesc1 = ReadCurrentLine.trim();
								}
							}

							// First Error message found
							if ((endOfResult == false)
									&& (MsgDesc1.toString().contentEquals(
											ReadCurrentLine.trim().toString()) == true)
									&& (WCAG2_Error1 == false)
									&& (MsgDesc1.toString().contentEquals("") == false)) {
								WCAG2_Error1 = true;
								lineFound = linenumber;
							} else {
								// Manipulate Error message 1
								if ((linenumber > lineFound)
										&& (WCAG2_Error1 == true)
										&& (endOfResult == false)
										&& (ReadCurrentLine.trim().toString()
												.contentEquals("") == false)
										&& (MsgDesc1.toString().contentEquals(
												ReadCurrentLine.trim()
														.toString()) == false)
										&& (MsgDesc1 == "" == false)) {
									// Error Message 1 Line 1 found
									if (linenumber == lineFound + 1) {
										if (ReadCurrentLine
												.contains("Column: ")) {
											// Extract the line number
											int Position = ReadCurrentLine
													.lastIndexOf("Column: ");
											char[] charArray = ReadCurrentLine
													.toCharArray();
											MsgDesc1Line1 = String.copyValueOf(
													charArray, 10,
													Position - 11);
											NumOfInstance1 = NumOfInstance1 + 1;
										}
									}
									// Error Message 1 Line 2 found
									if (linenumber == lineFound + 2) {
										ReadCurrentLine = ReadCurrentLine
												.replaceAll("href= ", "");
										MsgDesc1Line2 = MsgDesc1Line2
												+ "In line " + MsgDesc1Line1
												+ " on "
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

						if (MsgDesc1 == "" == false) {
							CommentDesc1 = MsgDesc1 + "\n" + MsgDesc1Line2
									+ "    Number of Instance: "
									+ NumOfInstance1;
							commentLine = NumOfInstance1 + 4;
						}
						WritableCellFeatures cellFeatures = new WritableCellFeatures();
						cellFeatures.setComment(CommentDesc1, 6, commentLine);
						Label label = new Label(Url_ToBeTest, 37, "Fail",
								arial9formatNoBold);
						label.setCellFeatures(cellFeatures);
						sheet2.addCell(label);
						WCAG2_Error1 = true;
						// End of Read Result file from WPSS tool

						// The 2.4.5 results Passed
						if (MsgDesc1 == "" == true) {
							// Result of Success Criterion WCAG2
							WritableCellFeatures cellFeatures2 = new WritableCellFeatures();
							// cellFeatures.setComment("Failed ",4,2);
							Label label2 = new Label(Url_ToBeTest, 37, "Pass",
									arial9formatNoBold);
							label2.setCellFeatures(cellFeatures2);
							sheet2.addCell(label2);
						}
					} // End of all URL While procedures
					catch (Exception ex) {
					}
				}

			}
		}
		// All cells modified/added. Now write out the workbook
		copy.write();
		copy.close();
		System.out
				.println("\n"
						+ "Congratulations!!!  The QA Review has been successfully completed.");
	}

	public static void main(String[] args) {
		new Editor();
	}

}
