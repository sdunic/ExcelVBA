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
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    utils.Unprotect Sheets(2)
    utils.Unprotect Sheets(3)
    globals.setAllowEventHandling False
        
    functions.loadChanges
    
    globals.setAllowEventHandling True
    utils.Protect Sheets(2)
    utils.Protect Sheets(3)
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    If globals.getAllowEventHandling = True Then
        utils.Unprotect Sheets(2)
        cfg.Init
        
        'NA
        If Not Intersect(Target, Range(cfg.getColNA_NovaCijena & ":" & cfg.getColNA_NovaCijena)) Is Nothing Then
            utils.setIndex Target.row, cfg.getColNA_Indeks, cfg.getColNA_Cijena, cfg.getColNA_NovaCijena, cfg.getColBrojPromjena
        End If
    
        'IA
        If Not Intersect(Target, Range(cfg.getColIA_NovaCijena & ":" & cfg.getColIA_NovaCijena)) Is Nothing Then
            utils.setIndex Target.row, cfg.getColIA_Indeks, cfg.getColIA_Cijena, cfg.getColIA_NovaCijena, cfg.getColBrojPromjena
        End If
        
        'Katalog
        If Not Intersect(Target, Range(cfg.getColKatalog_NovaCijena & ":" & cfg.getColKatalog_NovaCijena)) Is Nothing Then
            utils.setIndex Target.row, cfg.getColKatalog_Indeks, cfg.getColKatalog_Cijena, cfg.getColKatalog_NovaCijena, cfg.getColBrojPromjena
        End If
        
        'Rasprodaja
        If Not Intersect(Target, Range(cfg.getColRasprodaja_NovaCijena & ":" & cfg.getColRasprodaja_NovaCijena)) Is Nothing Then
            utils.setIndex Target.row, cfg.getColRasprodaja_Indeks, cfg.getColRasprodaja_Cijena, cfg.getColRasprodaja_NovaCijena, cfg.getColBrojPromjena
        End If
        
        'IstekRoka
        If Not Intersect(Target, Range(cfg.getColIstekRoka_NovaCijena & ":" & cfg.getColIstekRoka_NovaCijena)) Is Nothing Then
            utils.setIndex Target.row, cfg.getColIstekRoka_Indeks, cfg.getColIstekRoka_Cijena, cfg.getColIstekRoka_NovaCijena, cfg.getColBrojPromjena
        End If
        
        utils.Protect Sheets(2)
    End If
    
End Sub

