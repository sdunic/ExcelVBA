VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    utils.Unprotect Sheets(1)
    utils.Unprotect Sheets(2)
    utils.Unprotect Sheets(3)
    Sheets(1).Range("C2").Value = "v " & db.getDocVersion
    Sheets(2).Range("C2").Value = "v " & db.getDocVersion
    Sheets(3).Range("C2").Value = "v " & db.getDocVersion
    functions.insertLog "open_doc", "", ""
    cfg.Init
    
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    ActiveWorkbook.Sheets(1).Select
    Range("C5").Value = utils.getUserName()
       
    'poèetno stanje
    ActiveWorkbook.Sheets(1).Select
    Range("C7").Select
    Range("C13").ClearContents
    Range("C15").ClearContents
    
    utils.Protect Sheets(1)
    utils.Protect Sheets(2)
    utils.Protect Sheets(3)
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub
