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
    
    Range("B2").Value = "v " & CStr(db.getDocVersion)
    
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    ActiveWorkbook.Sheets(1).Select
    cfg.Init
    Range(cfg.get_korisnik & cfg.get_zaglavlje).Value = utils.getUserName
    If Len(Range(cfg.get_datumFakture & cfg.get_zaglavlje).Value) = 0 Then
        Range(cfg.get_datumFakture & cfg.get_zaglavlje).Value = Date
    End If
    Range(cfg.get_reasonCodeTekst & cfg.get_reasonCodeRedak).Activate
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub
