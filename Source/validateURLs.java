package Automation;

import java.io.File;
import java.io.IOException;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * <p>
 * <b>Reporting Tool Program: </b>Populates WPSS, CSE Pro and W3C results in the
 * MS-Excel Spreadsheet Report (Template)
 * </p>
 * <b>Description: </b> Permission is hereby granted, free of charge, to any
 * person obtaining a copy of this software and associated documentation files
 * (the "Software"). Therefore, the author reserve limitations and rights to
 * modify, merge, publish, sublicense and sell. Copyright has been reserved to
 * Matrixx Hi-Tech Inc. The Software can be distribute under the following
 * conditions:
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

public class validateURLs {

	private String inputFile;

	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public void read() throws IOException {
		boolean errorFound = false;
		File inputWorkbook = new File(inputFile);
		Workbook w;
		try {
			w = Workbook.getWorkbook(inputWorkbook);
			// Get the third sheet
			Sheet sheet = w.getSheet(2);
			// Loop over first 10 column and lines

			for (int j = 1; j < sheet.getColumns(); j++) {
				Cell cell = sheet.getCell(j, 5);
				if ((cell.getContents().length() > 6 == true)
						&& (cell.getContents().contains("[URL]") == false)) {
					if ((cell.getContents().contains("http:") == false)
							&& (cell.getContents().contains("https:") == false)) {
						errorFound = true;
						System.out.println(cell.getContents());
						break;
					}
				}
			}
		} catch (BiffException e) {
			e.printStackTrace();
		}
		if (errorFound == true) {
			// if the file didnt accept the renaming operation
			JOptionPane pane1 = new JOptionPane(
					"http or https protocol missing in one or all URL from input spreadsheet in row 5\n\n"
							+ "Please add http:// or https:// for all URL and try again\n");
			JDialog d1 = pane1.createDialog((JFrame) null,
					"http or https protocol missing in input spreadsheet");
			d1.setLocation(400, 500);
			d1.setVisible(true);
			System.exit(0);
		}
	}

	public static void main(String[] args) throws IOException {
		validateURLs test = new validateURLs();
		test.setInputFile("C:/QA-WS-Tool/Input.xls");
		test.read();
	}

}