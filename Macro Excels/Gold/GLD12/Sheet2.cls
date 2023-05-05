VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub insertOrders()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    utils.Unprotect Sheets(2)
    globals.setAllowEventHandling False
        
    functions.insertOrders
    
    globals.setAllowEventHandling True
    utils.Protect Sheets(2)
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub


