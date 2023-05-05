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
    
    functions.insertLog "open_doc", "", ""
    cfg.Init
    
    functions.checkVersion
    
    ActiveWorkbook.Sheets(1).Range("C2").Value = "v " & CStr(db.getDocVersion)
    ActiveWorkbook.Sheets(2).Range("C2").Value = "v " & CStr(db.getDocVersion)
    ActiveWorkbook.Sheets(3).Range("C2").Value = "v " & CStr(db.getDocVersion)
    
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    ActiveWorkbook.Sheets(1).Select
    Range("C5").Value = utils.getUserName()
       
    'poèetno stanje
    ActiveWorkbook.Sheets(1).Select
    Range("C7").Select
    Range("C9").ClearContents
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub

