; This script demonstrates using the Excel_Get function to get a reference to the active Excel application. Excel_Get 
; has a nice feature where it will exit edit-mode if you happen to be editing a cell when the function is called. 

F7::  ; Press F7 to display Excel's caption and the name of the active workbook.
    xlApp := Excel_Get()
    if !IsObject(xlApp)  ; If Excel_Get fails instead of returning an object it returns an error message.
    {
        MsgBox, 16, Excel_Get Error, % xlApp
        return
    }
    
    MsgBox, % "Caption: " xlApp.Caption "`n"
            . "Workbook: " xlApp.ActiveWorkbook.Name
return

Esc::ExitApp  ; Press Escape to exit this script.

; Excel_Get by jethrow (modified)
Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
    static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    WinGetClass, WinClass, %WinTitle%
    if !(WinClass == "XLMAIN")
        return "Window class mismatch."
    ControlGet, hwnd, hwnd,, Excel7%Excel7#%, %WinTitle%
    if (ErrorLevel)
        return "Error accessing the control hWnd."
    VarSetCapacity(IID_IDispatch, 16)
    NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
    if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc) != 0
        return "Error calling AccessibleObjectFromWindow."
    window := ComObject(9, pacc, 1)
    if !ComObjType(window) = 9
        return "Error wrapping the window object."
    Loop
        try return window.Application
        catch e
            if SubStr(e.message, 1, 10) = "0x80010001"
                ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
            else
                return "Error accessing the application object."
}  ; Needs testing

; References
;   https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
;   https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
