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
        'MPC A
        If Not Intersect(Target, Range(cfg.getColMPC_ANovaCijena & ":" & cfg.getColMPC_ANovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_AIndeks, cfg.getColMPC_ACijena, cfg.getColMPC_ANovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC B
        If Not Intersect(Target, Range(cfg.getColMPC_BNovaCijena & ":" & cfg.getColMPC_BNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_BIndeks, cfg.getColMPC_BCijena, cfg.getColMPC_BNovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC C
        If Not Intersect(Target, Range(cfg.getColMPC_CNovaCijena & ":" & cfg.getColMPC_CNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_CIndeks, cfg.getColMPC_CCijena, cfg.getColMPC_CNovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC D
        If Not Intersect(Target, Range(cfg.getColMPC_DNovaCijena & ":" & cfg.getColMPC_DNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_DIndeks, cfg.getColMPC_DCijena, cfg.getColMPC_DNovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC S1
        If Not Intersect(Target, Range(cfg.getColMPC_S1NovaCijena & ":" & cfg.getColMPC_S1NovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_S1Indeks, cfg.getColMPC_S1Cijena, cfg.getColMPC_S1NovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC S2
        If Not Intersect(Target, Range(cfg.getColMPC_S2NovaCijena & ":" & cfg.getColMPC_S2NovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_S2Indeks, cfg.getColMPC_S2Cijena, cfg.getColMPC_S2NovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC S3
        If Not Intersect(Target, Range(cfg.getColMPC_S3NovaCijena & ":" & cfg.getColMPC_S3NovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_S3Indeks, cfg.getColMPC_S3Cijena, cfg.getColMPC_S3NovaCijena, cfg.getColBrojPromjena
        End If
        
        'MPC KAMP
        If Not Intersect(Target, Range(cfg.getColMPC_KAMPNovaCijena & ":" & cfg.getColMPC_KAMPNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColMPC_KAMPIndeks, cfg.getColMPC_KAMPCijena, cfg.getColMPC_KAMPNovaCijena, cfg.getColBrojPromjena
        End If
    
    End If
    
End Sub

