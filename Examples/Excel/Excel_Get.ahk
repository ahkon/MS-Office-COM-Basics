; This script demonstrates using the Excel_Get function to get a reference to the active Excel application. Excel_Get 
; has a nice feature where it will exit edit-mode if you happen to be editing a cell when the function is called. 

F7::  ; Press F7 to display Excel's caption and the name of the active workbook.
    xlApp := Excel_Get()
    MsgBox, % "Caption: " xlApp.Caption "`n"
            . "Workbook: " xlApp.ActiveWorkbook.Name
return

/* ; Excel_Get by jethrow (modified)
 * Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
 *     static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
 *     WinGetClass, WinClass, %WinTitle%
 *     if (WinClass == "XLMAIN") {
 *         ControlGet, hwnd, hwnd, , Excel7%Excel7#%, %WinTitle%
 *         if !ErrorLevel {
 *             VarSetCapacity(IID, 16)
 *             NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID, "Int64"), "Int64")
 *             if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject, "Ptr", &IID, "Ptr*", pacc) = 0
 *                 window := ComObject(9, pacc, 1), ObjAddRef(pacc)
 *             if ComObjType(window) = 9
 *                 while !xl
 *                     try xl := window.Application
 *                     catch e
 *                         if SubStr(e.message, 1, 10) = "0x80010001"
 *                             ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
 *                         else
 *                             return "Error accessing the application object."
 *         }
 *     }
 *     return xl
 * }  ; Needs testing
 */
; Excel_Get by jethrow (modified)
Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
    WinGetClass, WinClass, %WinTitle%
    if (WinClass == "XLMAIN") {
        ControlGet, hwnd, hwnd, , Excel7%Excel7#%, %WinTitle%
        if !(ErrorLevel) {
            window := Acc_ObjectFromWindow(hwnd, -16)
            if ComObjType(window) = 9
                while !xl
                    try xl := window.application
                    catch e
                        if SubStr(e.message,1,10) = "0x80010001"
                            ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
                        else
                            return "Error accessing the application object."
        }
    }
    return xl
}

/* Acc_ObjectFromWindow(hWnd, idObject:=0) {
 *     Acc_Init()
 *     idObject &= 0xFFFFFFFF
 *     VarSetCapacity(IID, 16)
 *     if (idObject == 0xFFFFFFF0)
 *         NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID, "Int64"), "Int64")
 *     else
 *         NumPut(0x719B3800AA000C81, NumPut(0x11CF3C3D618736E0, IID, "Int64"), "Int64")
 * 
 *     If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject, "Ptr", &IID, "Ptr*", pacc) = 0
 *         return ComObject(9, pacc, 1), ObjAddRef(pacc)
 * }
 */
Acc_ObjectFromWindow(hWnd, idObject:=0) {
    Acc_Init()
    If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
        return ComObject(9, pacc, 1), ObjAddRef(pacc)
}

Acc_Init() {
    static h
    If !h
        h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}

; References
;   https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
;   https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
