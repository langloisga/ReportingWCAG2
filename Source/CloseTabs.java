package Automation;

//BEGIN custom imports 	
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;

/**
 * <b>Java Program LaunchFirefox</b> <b>Description: </b> Launch Firefox and
 * initialize Browser
 * 
 * 
 * @author Gaston Langlois - Environment Canada
 * @since 2012/06/20
 */

public class CloseTabs {
	private static String PageTitle = null;

	public static void main(String [] args) throws AWTException {
		Robot robot = new Robot();

		////////////////// MAKE SURE FIREFOX IS YOUR DEFAULT BROWSER ///////////////
		// Will Close tab(s) in Browser Firefox 
		// Right-Click on the first tab in Firefox
		robot.mouseMove(200, 55);
		robot.mousePress(InputEvent.BUTTON3_MASK);
		robot.mouseRelease(InputEvent.BUTTON3_MASK);
		robot.delay(800);
		// Press Ctrl+C to Close tab
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_O);
		robot.keyRelease(KeyEvent.VK_O);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.delay(500);

		// Checking If there are more than one tab opened to close
		HWND hwnd = User32.INSTANCE.FindWindow(null, "Confirm close");
		if (hwnd != null) {
			// System.out.println("Windows found");
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);

		}
		// TODO Get the current page from Mozilla Firefox
		PageTitle= "test";
		// else {
		// //System.out.println("Windows not found");
		//
		// }
	}

	public static String getPagetitle() {
		return PageTitle;
	}

}