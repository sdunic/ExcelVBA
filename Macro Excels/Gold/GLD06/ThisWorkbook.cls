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
    
    cfg.Init
    
    'dohvat domenskog korisnika prilikom otvaranja dokumenta
    ActiveWorkbook.Sheets(1).Select
    Range("C5").Value = utils.getUserName()
        
    'poèetno stanje
    ActiveWorkbook.Sheets(1).Select
    Range("C7").Select
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
