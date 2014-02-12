package Automation;

//BEGIN custom imports 		

import com.sun.jna.Native;
import com.sun.jna.PointerType;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.RECT;
import com.sun.jna.win32.W32APIOptions;

/**
 * <b>Java Program</b> <b>Description: </b> Check browser
 * 
 * 
 * @author Gaston Langlois - Environment Canada
 * @since 2012/06/20
 */

public class MinimizeAllWindow {
	public interface User32 extends W32APIOptions {
		public static final String SHELL_TRAY_WND = "Shell_TrayWnd";
		public static final int WM_COMMAND = 0x111;
		public static final int MIN_ALL = 0x1a3;
		public static final int MIN_ALL_UNDO = 0x1a0;
		User32 INSTANCE = (User32) Native.loadLibrary("user32", User32.class);
		HWND GetForegroundWindow(); // add this
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

	public static void main(String[] args) {
		// get the taskbar's window handle
		HWND shellTrayHwnd = User32.instance.FindWindow(User32.SHELL_TRAY_WND,
				null);

		// use it to minimize all windows
		User32.instance.SendMessageA(shellTrayHwnd, User32.WM_COMMAND,
				User32.MIN_ALL, 0);

		// Test
		//System.out.println(shellTrayHwnd);
		//System.out.println(User32.WM_COMMAND);

		// sleep for 3 seconds
		try {
			Thread.sleep(3000);
		} catch (InterruptedException e) {
		}

		// then restore previously minimized windows
		User32.instance.SendMessageA(shellTrayHwnd, User32.WM_COMMAND, User32.MIN_ALL_UNDO, 0);
		byte[] windowText = new byte[512];
		PointerType hwnd = User32.INSTANCE.GetForegroundWindow();
		// then you can call it!
		User32.INSTANCE.GetWindowTextA(hwnd, windowText, 512);
		System.out.println(Native.toString(windowText));
		//System.out.println(User32.INSTANCE.GetForegroundWindow());

	}
}