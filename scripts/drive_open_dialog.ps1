# Drives the native Windows "Open" file dialog (class #32770) without needing
# focus: writes the path into the dialog's filename Edit control via WM_SETTEXT
# and presses its Open button via BM_CLICK. Used by the Google Sheets sync
# workflow (File -> Import -> Upload -> Browse in the Finance spreadsheet) so
# the whole import can run unattended after the browser clicks Browse.
#
# Why not the obvious approaches (both proven to FAIL on this machine, 2026-07-11):
#   - $wshell.AppActivate('Open') + SendKeys: AppActivate matches loosely and the
#     keystrokes once landed in the wrong window (typed the path straight into a
#     terminal). Never use blind SendKeys for this.
#   - SetForegroundWindow: blocked by Windows foreground-lock; a background
#     process cannot steal focus, so focus-then-type cannot be made reliable.
# SendMessage directly to the dialog's child controls needs no focus at all.
#
# Usage: powershell -File drive_open_dialog.ps1 -Path "C:\Users\Paul\Downloads\Stocks_Buy_Strategy.xlsx"
#   Exits 0 on success, 1 if no Open dialog was found or its controls were missing.
#   The dialog must already be open (click the page's Browse/Choose-file button first).
param(
    [Parameter(Mandatory = $true)][string]$Path,
    # Window title of the file dialog. Chrome's file picker is titled "Open".
    [string]$DialogTitle = 'Open',
    # How long to keep looking for the dialog before giving up, in seconds.
    [int]$TimeoutSeconds = 10
)

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class OpenDialogDriver {
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lp);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr lp);
  [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, string lp);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
  delegate bool EnumProc(IntPtr h, IntPtr lp);
  const uint WM_SETTEXT = 0x000C;
  const uint BM_CLICK = 0x00F5;

  public static IntPtr FindDialog(string title) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, lp) => {
      if (!IsWindowVisible(h)) return true;
      var c = new StringBuilder(256); GetClassName(h, c, 256);
      var t = new StringBuilder(256); GetWindowText(h, t, 256);
      if (c.ToString() == "#32770" && t.ToString() == title) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }

  public static string Drive(IntPtr dialog, string path) {
    IntPtr edit = IntPtr.Zero, openBtn = IntPtr.Zero;
    EnumChildWindows(dialog, (h, lp) => {
      var c = new StringBuilder(256); GetClassName(h, c, 256);
      var t = new StringBuilder(256); GetWindowText(h, t, 256);
      if (c.ToString() == "Edit" && edit == IntPtr.Zero) edit = h;
      if (c.ToString() == "Button" && t.ToString().Replace("&", "") == "Open") openBtn = h;
      return true;
    }, IntPtr.Zero);
    if (edit == IntPtr.Zero) return "FAILED: no Edit control found in dialog";
    if (openBtn == IntPtr.Zero) return "FAILED: no Open button found in dialog";
    SendMessage(edit, WM_SETTEXT, IntPtr.Zero, path);
    System.Threading.Thread.Sleep(400);
    SendMessage(openBtn, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
    return "SUCCESS";
  }
}
"@

if (-not (Test-Path $Path)) {
    Write-Output "FAILED: file does not exist: $Path"
    exit 1
}

$dialog = [IntPtr]::Zero
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
    $dialog = [OpenDialogDriver]::FindDialog($DialogTitle)
    if ($dialog -ne [IntPtr]::Zero) { break }
    Start-Sleep -Milliseconds 500
}
if ($dialog -eq [IntPtr]::Zero) {
    Write-Output "FAILED: no visible '$DialogTitle' dialog (class #32770) found within $TimeoutSeconds s"
    exit 1
}

$result = [OpenDialogDriver]::Drive($dialog, $Path)
Write-Output "$result (dialog hwnd $dialog)"
if ($result -ne 'SUCCESS') { exit 1 }
