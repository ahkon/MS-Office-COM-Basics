FilePath := A_ScriptDir "\New Microsoft Word Document.docx"  ; Path to a Word document.
if FileOpen(FilePath, "rw") ; FileOpen fails if the file is already open.
    MsgBox, 64, File Status, % "The file is not open.`n`n(" FilePath ")"
else
    MsgBox, 64, File Status, % "The file is already open.`n`n(" FilePath ")"
