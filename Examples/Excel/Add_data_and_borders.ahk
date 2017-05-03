; This script copies data from an example object into Excel.

; Constants
msoTrue := -1
xlContinuous := 1
xlDown := -4121
xlEdgeBottom := 9
xlEdgeLeft := 7
xlEdgeRight := 10
xlEdgeTop := 8
xlInsideHorizontal := 12
xlInsideVertical := 11
xlMedium := -4138
xlRight := -4152

; The example object. 
ExObj := [ {Quantity: 2, Item: "Abc", Code: "asdf1234"}
         , {Quantity: 3, Item: "Xyz", Code: "zxcv2345"} 
         , {Quantity: 5, Item: "Lmn", Code: "qwer3456"}
         , {Quantity: 7, Item: "Opq", Code: "dfgh4567"} ]

; Open Excel, input the list of items and format for printing.
xlApp := ComObjCreate("Excel.Application")  ; Create an Excel Application object and save a reference to it.
WrkBk := xlApp.Workbooks.Add                ; Create a new workbook object and save a reference to it.

; Create a SafeArray containing the proper amount of rows and columns. 4 extra rows are added for the column headings,
; a blank row, user and date.
SafeArray := ComObjArray(12, ExObj.MaxIndex() + 4, 3)

; Headings
SafeArray[0, 0] := "User:", SafeArray[0, 1] := A_UserName
SafeArray[1, 0] := "Date:"

SafeArray[3, 0] := "Quantity"
SafeArray[3, 1] := "Item"
SafeArray[3, 2] := "Code"

; Insert the items from ExObj into the SafeArray.
for RowNumber, Row in ExObj
    for FieldName, FieldValue in Row
        SafeArray[RowNumber + 3, A_Index - 1] := FieldValue

TopLeftCell := xlApp.Worksheets(1).Cells(1, 1)  ; The top left cell where the data will be inserted.
BotRightCell := xlApp.Worksheets(1).Cells(ExObj.MaxIndex() + 4, 3)  ; Bot. right cell where the data will be inserted.
TotalRange := xlApp.Range(TopLeftCell, BotRightCell)
TotalRange.Value := SafeArray  ; Copy the SafeArray into the range.

xlApp.Worksheets(1).Range("B2").NumberFormat := "@"	; NumberFormat @=Text
xlApp.Worksheets(1).Range("B2").Value := A_MMMM " " A_DD ", " A_YYYY

; Format: Borders, Bold
ThisRange := xlApp.Worksheets(1).Range("A1:A2")
ThisRange.Font.Bold := msoTrue
ThisRange.HorizontalAlignment := xlRight

ThisRange := xlApp.Worksheets(1).Range("A4:C4")
ThisRange.Font.Bold := msoTrue
ThisRange.Borders(xlInsideVertical).LineStyle := xlContinuous
ThisRange.Borders(xlInsideVertical).Weight := xlMedium
for i, Const in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight] {
    ThisRange.Borders(Const).LineStyle := xlContinuous
    ThisRange.Borders(Const).Weight := xlMedium
}

; Set column widths.
for i, Col in [1, 3] {
    TopLeftCell := xlApp.Worksheets(1).Cells(4, Col)  ; Do not include rows 1-3.
    BotRightCell := TopLeftCell.End(xlDown)
    xlApp.Worksheets(1).Range(TopLeftCell, BotRightCell).Columns.AutoFit
}
xlApp.Worksheets(1).Columns("B:B").ColumnWidth := 40  ; Fixed width column.

TopLeftCell := xlApp.Worksheets(1).Cells(5, 1)
BotRightCell := xlApp.Worksheets(1).Cells(ExObj.MaxIndex() + 4, 3)
ThisRange := xlApp.Range(TopLeftCell, BotRightCell)
ThisRange.Borders(xlInsideVertical).LineStyle := xlContinuous
ThisRange.Borders(xlInsideVertical).Weight := xlMedium
ThisRange.Borders(xlInsideHorizontal).LineStyle := xlContinuous
ThisRange.Borders(xlInsideHorizontal).Weight := xlThin := 2
for i, Const in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight] {
    ThisRange.Borders(Const).LineStyle := xlContinuous
    ThisRange.Borders(Const).Weight := xlMedium
}

xlApp.Visible := true
TotalRange.Select
WrkBk.Saved := msoTrue
WrkBk.PrintPreview(msoTrue)
