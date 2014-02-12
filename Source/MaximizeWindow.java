package Automation;
		
//BEGIN custom imports 		
import com.sun.jna.platform.win32.WinDef.*;
import com.sun.jna.platform.win32.User32;


/**
 * <b>Java Program</b> <b>Description: </b> Check browser
 * 
 * @author Gaston Langlois - Environment Canada
 * @since 2012/06/20
 */

public class MaximizeWindow {
	
	public static void main(String[] args) {
		HWND hwnd = User32.INSTANCE.FindWindow
		     //  (null, "TEST.txt - Notepad"); // window title
				 (null, "Confirm close");
		if (hwnd != null) {
			System.out.println("Windows found");

		}
		else {
			System.out.println("Windows not found");
			
		}
		
	}
}