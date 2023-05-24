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
    
        'TNC
        If Not Intersect(Target, Range(cfg.getColTNC_NovaCijena & ":" & cfg.getColTNC_NovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_Indeks, cfg.getColTNC_Cijena, cfg.getColTNC_NovaCijena, cfg.getColBrojPromjena
        End If
    
        'TNC A
        If Not Intersect(Target, Range(cfg.getColTNC_ANovaCijena & ":" & cfg.getColTNC_ANovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_AIndeks, cfg.getColTNC_ACijena, cfg.getColTNC_ANovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC B
        If Not Intersect(Target, Range(cfg.getColTNC_BNovaCijena & ":" & cfg.getColTNC_BNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_BIndeks, cfg.getColTNC_BCijena, cfg.getColTNC_BNovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC C
        If Not Intersect(Target, Range(cfg.getColTNC_CNovaCijena & ":" & cfg.getColTNC_CNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_CIndeks, cfg.getColTNC_CCijena, cfg.getColTNC_CNovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC D
        If Not Intersect(Target, Range(cfg.getColTNC_DNovaCijena & ":" & cfg.getColTNC_DNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_DIndeks, cfg.getColTNC_DCijena, cfg.getColTNC_DNovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC S
        If Not Intersect(Target, Range(cfg.getColTNC_SNovaCijena & ":" & cfg.getColTNC_SNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_SIndeks, cfg.getColTNC_SCijena, cfg.getColTNC_SNovaCijena, cfg.getColBrojPromjena
        End If
        
        'TNC KAMP
        If Not Intersect(Target, Range(cfg.getColTNC_KAMPNovaCijena & ":" & cfg.getColTNC_KAMPNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColTNC_KAMPIndeks, cfg.getColTNC_KAMPCijena, cfg.getColTNC_KAMPNovaCijena, cfg.getColBrojPromjena
        End If
    
    End If
    
End Sub

