; Get a reference a cell in the active Excel application.
xlApp := ComObjActive("Excel.Application")  ; Excel must be running.

MyCell := xlApp.Cells(2, 3)  ; Get the cell at row 2, column 3.
; MyCell := xlApp.Range("C2")  ; Get the cell "C2". (This does the same thing as the previous line.)

; Display the results.
MsgBox, % "The cell at address " MyCell.Address " has a value of '" MyCell.Value 
        . "' and has the text '" MyCell.Text "'."
