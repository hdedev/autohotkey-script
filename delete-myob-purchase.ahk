#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode 2

^Esc::ExitApp
F9::Pause -1

; --- CONFIG ---
excelPath := "D:\temp\delete3.xlsx"
sheetName := "Sheet1"
col := 1          ; Column A
startRow := 1
delayUI := 800

; --- OPEN EXCEL ---
xl := ComObject("Excel.Application")
xl.Visible := false
wb := xl.Workbooks.Open(excelPath)
ws := wb.Worksheets(sheetName)

row := startRow

; Activate MYOB
WinActivate "Cozy Australia - MYOB AccountRight"
WinWaitActive "Cozy Australia - MYOB AccountRight"

Loop {
    code := ws.Cells(row, col).Value
    if (code = "")
        break

	if (row > 55)
		break
		

    ; Ctrl + Shift + F (search)
    Send "^+f"
    Sleep 200
	
	SendText "`""
    SendText code
	SendText "`""
    Sleep 200
    Send "{Enter}"

    Sleep delayUI

	Send "{Tab}"
	Send "{Down}"
	Send "{Enter}"
	
	Sleep delayUI
	
	
	
	result := MsgBox("Delete this purchase?"  row code, "Confirm", "YesNo")

	if (result = "No"){
		Send "{Esc}"
		Sleep 200
		row := row + 1
		continue   ; skip to next record
	}
	Send "!e"
	Sleep 200
	Send "e"
	
	Sleep delayUI
	Sleep delayUI
    
    row := row + 1 
}

wb.Close(false)
xl.Quit()

MsgBox "Done!"