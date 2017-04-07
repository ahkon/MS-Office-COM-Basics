; This script is a demonstration of the .Find method. It gets the value of the cell to the right of the active cell.
; Then it will find that value in the range "B:B" on Sheet2.

; Constants
xlValues := -4163
xlWhole := 1

xlApp := ComObjActive("Excel.Application")  ; Get a reference to the running instance of Excel.

; Get the value to the right of the active cell. This is the value that we will search for on the other sheet.
cellValue := xlApp.ActiveCell.Offset(0, 1).Value

; Save a reference to "Sheet2". Specify the sheet name (ie: "Sheet2") or number (ie: 2).
xlSht1 := xlApp.ActiveWorkbook.Sheets(2)

FoundCell := xlSht1.Range("B:B").Find(cellValue)

if (FoundCell)
{
    MsgBox, 64, Found, % FoundCell.Address
}
else
{
    MsgBox, 64, Not Found, This code does not exist
}
return

; References
; https://autohotkey.com/boards/viewtopic.php?f=5&t=30239
