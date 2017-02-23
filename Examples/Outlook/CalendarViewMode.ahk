; This script activates the calendar and changes the view to the month calendar.

; Constants
olFolderCalendar := 9
olCalendarViewMonth := 2

objOL := ComObjCreate("Outlook.Application")  ; Create an application object.
objExplorer := objOL.ActiveExplorer  ; Get a reference to the explorer window.
objExplorer.CurrentFolder := objOL.Session.GetDefaultFolder(olFolderCalendar)  ; Switch to the calendar folder.

; Apply view. This is required, for example, if the current callendar view is "List."
; If you try to set CalendarViewMode without a calendar view it will throw an error.
objExplorer.CurrentFolder.Views.Item("Calendar").Apply

objExplorer.CurrentView.CalendarViewMode := olCalendarViewMonth
objExplorer.CurrentView.Save

; References
;   https://autohotkey.com/boards/viewtopic.php?f=5&t=21687&p=104558#p104558
