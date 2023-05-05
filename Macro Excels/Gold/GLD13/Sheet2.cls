VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub setNewPricesAndIndexes()
    globals.setAllowEventHandling False
    
    'tu æe biti klik koji æe pokrenuti funkciju za MPC politiku za A, B, C, D i S.
    
    globals.setAllowEventHandling True
End Sub

Sub loadPriceChanges()
    globals.setAllowEventHandling False
        
    functions.loadChanges
    
    globals.setAllowEventHandling True
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    
    If globals.getAllowEventHandling = True Then
        'MPC KAMP
        If Not Intersect(Target, Range(cfg.getColMPC_KAMPNovaCijena & ":" & cfg.getColMPC_KAMPNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_KAMPIndeks, cfg.getColMPC_KAMPCijena, cfg.getColMPC_KAMPNovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC KAMP
        If Not Intersect(Target, Range(cfg.getColTNC_KAMPNovaCijena & ":" & cfg.getColTNC_KAMPNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_KAMPIndeks, cfg.getColTNC_KAMPCijena, cfg.getColTNC_KAMPNovaCijena, cfg.getColBrojPromjena
        End If
    
    End If
    
End Sub

