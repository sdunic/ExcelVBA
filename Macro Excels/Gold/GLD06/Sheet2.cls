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
        
        'Konzum Hiper
        If Not Intersect(Target, Range(cfg.getColKonzumHiperNovaCijena & ":" & cfg.getColKonzumHiperNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColKonzumHiperIndeks, cfg.getColKonzumHiperCijena, cfg.getColKonzumHiperNovaCijena, cfg.getColBrojPromjena
        End If
        
        'Konzum Maxi
        If Not Intersect(Target, Range(cfg.getColKonzumMaxiNovaCijena & ":" & cfg.getColKonzumMaxiNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColKonzumMaxiIndeks, cfg.getColKonzumMaxiCijena, cfg.getColKonzumMaxiNovaCijena, cfg.getColBrojPromjena
        End If
        
        'Studenac
        If Not Intersect(Target, Range(cfg.getColStudenacNovaCijena & ":" & cfg.getColStudenacNovaCijena)) Is Nothing Then
            cfg.Init
            utils.setIndex Target.row, cfg.getColStudenacIndeks, cfg.getColStudenacCijena, cfg.getColStudenacNovaCijena, cfg.getColBrojPromjena
        End If
        
        
    
    End If
    
End Sub

