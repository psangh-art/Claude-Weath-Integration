# Bring the TradingView Desktop main window to the FOREGROUND (user request
# 2026-07-20: selecting a layout from any dashboard screen must present the
# TradingView window, not leave it navigating in the background).
#
# A bare SetForegroundWindow from a background process is blocked by Windows'
# foreground-lock (see CLAUDE.md, drive_open_dialog.ps1 note). The reliable
# work-around is: un-minimise via ShowWindow(SW_RESTORE), tap ALT to reset the
# foreground-lock timer, then AttachThreadInput to the current foreground
# window's thread so SetForegroundWindow is permitted. Same no-window-focus
# user32 family already used by ensureWindowMaximized / drive_open_dialog.ps1.
#
# Exit 0 if a TradingView window was found and raised, 1 if none was found.
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Fg {
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pid);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] public static extern void keybd_event(byte v, byte s, uint f, IntPtr e);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  public static bool Show(IntPtr h) {
    if (IsIconic(h)) { ShowWindow(h, 9); } else { ShowWindow(h, 5); } // SW_RESTORE / SW_SHOW
    // A synthetic ALT tap resets the foreground-lock timeout so the call below is allowed.
    keybd_event(0x12, 0, 0, IntPtr.Zero);
    keybd_event(0x12, 0, 2, IntPtr.Zero);
    IntPtr fg = GetForegroundWindow();
    uint fgThread = GetWindowThreadProcessId(fg, IntPtr.Zero);
    uint cur = GetCurrentThreadId();
    AttachThreadInput(cur, fgThread, true);
    BringWindowToTop(h);
    bool ok = SetForegroundWindow(h);
    AttachThreadInput(cur, fgThread, false);
    return ok;
  }
}
'@
$hit = $false
Get-Process TradingView -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object {
  [Fg]::Show($_.MainWindowHandle) | Out-Null
  $hit = $true
}
if (-not $hit) { exit 1 }
