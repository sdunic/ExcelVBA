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
    
    functions.insertLog "open_doc", "", ""
    
    cfg.Init
    ActiveWorkbook.Sheets(1).Select
    Range("B2").Value = "v " & CStr(db.getDocVersion)
    Range(cfg.get_reasonCodeTekst & cfg.get_reasonCodeRedak).Activate
    Range(cfg.get_korisnik & cfg.get_zaglavlje).Value = utils.getUserName
    If Len(Range(cfg.get_datumFakture & cfg.get_zaglavlje).Value) = 0 Then
        Range(cfg.get_datumFakture & cfg.get_zaglavlje).Value = Date
    End If
    
    ActiveWorkbook.Sheets(2).Select
    Range("B2").Value = "v " & CStr(db.getDocVersion)
    Range(cfg.get_reasonCodeTekst & cfg.get_reasonCodeRedak).Activate
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    Range(cfg.get_korisnik & cfg.get_zaglavlje).Value = utils.getUserName
    
    
    ActiveWorkbook.Sheets(1).Select
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub
