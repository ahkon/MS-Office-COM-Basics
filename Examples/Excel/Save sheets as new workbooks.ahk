; This script prompts the user to select a workbook file. Then each sheet in the selected workbook is copied into a
; new workbook. Each new workbook is then saved to the same dir as the original file.

FileSelectFile, SelectedFile, 3, %A_ScriptDir%, Select a file, Excel Workbooks (*.xls; *.xlsx)
if !FileExist(SelectedFile)
    return
SplitPath, SelectedFile, SelectedFileName, SecectedFileDir, SelectedFileExtension, SelectedNameNoExt
xlApp := ComObjCreate("Excel.Application")
xlApp.Visible := true
wbkSource := xlApp.Workbooks.Open(SelectedFile)
for sht, in wbkSource.Sheets
{
    wb := xlApp.Workbooks.Add()
    sht.Copy(wb.Sheets(1))
    wb.SaveAs(FindFreeName(SecectedFileDir, sht.Name, SelectedFileExtension))
    wb.Close(0)
}
wbkSource.Close(0)
xlApp.Quit
return

; Find an unused file name.
FindFreeName(FileDir, FileName, FileExt)
{
    FilePath := FileDir "\" FileName "." FileExt
    if FileExist(FilePath)  ; If the file exists, loop until an unused name is found.
    {
        Loop,
        {
            FilePath := FileDir "\" FileName "(" (A_Index + 1) ")." FileExt
            if !FileExist(FilePath)
                break
        }
    }
    return, FilePath
}
