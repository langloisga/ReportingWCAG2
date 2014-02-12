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

public class LaunchFirefox {
	private static String PageTitle = null;

	public static void main(String [] args) throws AWTException {
		String url = "http://ec.gc.ca";
		String os = System.getProperty("os.name").toLowerCase();
		Runtime rt = Runtime.getRuntime();
		Robot robot = new Robot();

		try {

			if (os.indexOf("win") >= 0) {

				// this doesn't support showing urls in the form of
				// "page.html#nameLink"
				rt.exec("rundll32 url.dll,FileProtocolHandler " + url);

			} else if (os.indexOf("mac") >= 0) {

				rt.exec("open " + url);

			} else if (os.indexOf("nix") >= 0 || os.indexOf("nux") >= 0) {

				// Do a best guess on unix until we get a platform independent
				// way
				// Build a list of browsers to try, in this order.
				String[] browsers = { "epiphany", "firefox", "mozilla",
						"konqueror", "netscape", "opera", "links", "lynx" };

				// Build a command string which looks like
				// "browser1 "url" || browser2 "url" ||..."
				StringBuffer cmd = new StringBuffer();
				for (int i = 0; i < browsers.length; i++)
					cmd.append((i == 0 ? "" : " || ") + browsers[i] + " \""
							+ url + "\" ");

				rt.exec(new String[] { "sh", "-c", cmd.toString() });

			} else {
				// return;
			}
		} catch (Exception e) {

		}
		robot.delay(2000);
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