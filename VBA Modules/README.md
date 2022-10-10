# VBA Tool snippets
This folder contains a collection of single-purpose VBA modules.

Note: All modules are made for MS Excel, although some might work with other applications. YMMV.

## List
List and description of the included modules.

| File module | Description | Usage |
| ----- | ----- | ----- |
| *toolCopyWorkbookModule.bas* | Creates a temporary copy of the current workbook | Call Function `CopyWorkbook()` <br/>Returns Workbook object of the copy<br/>Use optional boolean parameters `closeOriginal` (close source file) and `activateNew` (set new file to foreground) - both are `False` by default |
| *toolSleep.bas* | Provides a sleep function via kernel32.dll | Call `Sleep(ms)` with Long `ms` as milliseconds<br/>See **Examples/splashScreen.xlsm** for a demonstration |
| *FrmWait.frm* | Splash screen for visualizing background tasks | Call `FrmWait.SetText(str)` to display the windows with text `str`, call it again to change the text (such as updating a percentage)<br/>Call `FrmWait.Remove` to remove the window again.<br/>See **Examples/splashScreen.xlsm** for a demonstration |
| *toolSystemTickCount.bas* | Gets System TickCount i.e. time in milliseconds since last system boot via kernel32.dll. Can be used for differential timers. | Call `GetSystemTickCount()`, returns Long<br/>See **Examples/getSystemTickCount.xlsm**
| *toolMd5sum.bas* | Enables MD5 hash sums in Excel makros and formulas. Works over multiple cells. | Call `=StringToMD5Hex(...)` to hash single strings. <br/>Call `=Md5Hash(...)` to hash single cell or range.