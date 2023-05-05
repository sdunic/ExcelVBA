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
    ActiveWorkbook.Sheets(1).Select
    Range("C5").Value = utils.getUserName()
    Range("C8").Select
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    functions.insertLog "close_doc", "", ""
End Sub
