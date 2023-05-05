VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub

Private Sub Workbook_Open()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    Sheets(1).Unprotect
    
    Sheets(1).Range("C2").Value = "v " & CStr(Replace(db.getDocVersion, ",", "."))
    functions.insertLog "open_doc", "", ""
        
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    ActiveWorkbook.Sheets(1).Select
    Range("C5").Value = utils.getUserName()
    
    'poèetno stanje i postavljanje padajuæeg izbornika
    ActiveWorkbook.Sheets(1).Select

    Range("C7").Select
    
    
    Sheets(1).Protect
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
