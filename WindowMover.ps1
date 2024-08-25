Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WindowAPI {
	[DllImport("user32.dll")] //we have to import this Dynamic link library as this contains methods for getting and setting window locations.
	[return: MarshalAs(UnmanagedType.Bool)]
	public static extern bool GetWindowRect( //Used to get Window coordinates
		IntPtr hWnd, out RECT lpRect);
		
	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	public extern static bool MoveWindow( //Used to move windows
		IntPtr handle, int x, int y, int width, int height, bool redraw);
		
	[DllImport("user32.dll")]
	[return: MarshalAs(UnmanagedType.Bool)]
	public static extern bool SetForegroundWindow(IntPtr hWnd); //Used to bring window to foreground :)
}
public struct RECT {
	public int Left;        // x position of upper-left corner
	public int Top;         // y position of upper-left corner
	public int Right;       // x position of lower-right corner
	public int Bottom;      // y position of lower-right corner
}
"@
