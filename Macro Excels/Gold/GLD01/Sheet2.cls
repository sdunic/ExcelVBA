VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub loadChanges()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
        
    cfg.Init
    functions.loadConditionChanges
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If globals.getAllowEventHandling = True Then
        cfg.Init
        If Target.row > 5 And Target.column >= Range(cfg.getcTNUPACH & Target.row).column And Target.column <= Range(cfg.getcTNUDFIN606 & Target.row).column Then
            
            Application.Cursor = xlWait
            Application.ScreenUpdating = False
            Dim redak As String
            redak = Range("C" & Target.row).Value
            Dim oldTnunnc As String
            ActiveWorkbook.Sheets(4).Visible = xlSheetVisible
            ActiveWorkbook.Sheets(4).Select
                                    
            Dim i, j As Long
            For i = Target.column To (Target.column + (Target.Count / Target.rows.Count) - 1)
                For j = Target.row To (Target.row + Target.rows.Count - 1)
                    Dim cell As Range
                    
                    If Target.row > 5 And (Target.column = Range(cfg.getcTNUNNC & Target.row).column Or _
                                           Target.column = Range(cfg.getcTNUPADDEB & Target.row).column Or _
                                           Target.column = Range(cfg.getcTNUPADFIN & Target.row).column) Then
                        ActiveWorkbook.Sheets(4).Range("A:A").Select
                        Set cell = Selection.Find(What:=Cells(Target.row, Target.column).address, After:=ActiveCell, LookIn:=xlFormulas, _
                            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                            MatchCase:=False, SearchFormat:=False)
                                           
                        oldTnunnc = ActiveWorkbook.Sheets(4).Range("D" & cell.row).Value
                    End If
                    
                    Dim cellLocation As String
                    cellLocation = Cells(j, i).address
                
                    
                    ActiveWorkbook.Sheets(4).Range("A:A").Select
                    Set cell = Selection.Find(What:=cellLocation, After:=ActiveCell, LookIn:=xlFormulas, _
                            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                            MatchCase:=False, SearchFormat:=False)
                            
                    Dim targetValue As String
                    Dim bpr As String
                    If Target.Count = 1 Then
                        targetValue = Target.Value
                    Else
                        targetValue = Target.Value2(j - Target.row + 1, i - Target.column + 1)
                    End If
                                        
                    If Target.column = Range(cfg.getcTNUPADDEB & Target.row).column Or Target.column = Range(cfg.getcTNUPADFIN & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB601 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN601 & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB602 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN602 & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB603 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN603 & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB604 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN604 & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB605 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN605 & Target.row).column _
                       Or Target.column = Range(cfg.getcTNUDDEB606 & Target.row).column Or Target.column = Range(cfg.getcTNUDFIN606 & Target.row).column _
                    Then
                        targetValue = utils.getDateFormat(targetValue)
                    ElseIf Target.column = Range(cfg.getcTNUVAL601 & Target.row).column Or _
                            Target.column = Range(cfg.getcTNUVAL602 & Target.row).column Or _
                            Target.column = Range(cfg.getcTNUVAL603 & Target.row).column Or _
                            Target.column = Range(cfg.getcTNUVAL604 & Target.row).column Or _
                            Target.column = Range(cfg.getcTNUVAL605 & Target.row).column Or _
                            Target.column = Range(cfg.getcTNUVAL606 & Target.row).column _
                    Then
                        targetValue = Format(targetValue, "#.00")
                    End If
                    
                    If cell Is Nothing Then
                        MsgBox "Greška broj 001: Nedostaje inicijalni podatak!", vbOKOnly, "Greška"
                    Else
                        
                        If Len(targetValue) > 0 And targetValue <> CStr(ActiveWorkbook.Sheets(4).Range("C" & cell.row).Value2) Then
                            ActiveWorkbook.Sheets(4).Range("D" & cell.row).Value = targetValue
                            ActiveWorkbook.Sheets(4).Range("E" & cell.row).Value = 1
                            ActiveWorkbook.Sheets(4).Range("F" & cell.row).Value = utils.getUserName
                            
                            If Target.Font.ColorIndex <> 3 Then
                                bpr = cfg.getcBROJPROMJENA
                                Range(bpr & j).Value = Range(bpr & j).Value + 1
                            End If
                            
                            With Cells(j, i).Font
                                .Color = -16776961
                                .TintAndShade = 0
                            End With
                        Else
                            ActiveWorkbook.Sheets(4).Range("D" & cell.row).ClearContents
                            ActiveWorkbook.Sheets(4).Range("E" & cell.row).ClearContents
                            ActiveWorkbook.Sheets(4).Range("F" & cell.row).ClearContents
                            With Cells(j, i).Font
                                .Color = -1
                                .TintAndShade = 0
                            End With
                            bpr = cfg.getcBROJPROMJENA
                            Range(bpr & j).Value = Range(bpr & j).Value - 1
                            If Range(bpr & j).Value < 1 Then
                                Range(bpr & j).ClearContents
                            End If
                        End If
                    End If
            
                Next j
            Next i
            
            If Target.row > 5 And (Target.column = Range(cfg.getcTNUNNC & Target.row).column Or _
                                   Target.column = Range(cfg.getcTNUPADDEB & Target.row).column Or _
                                   Target.column = Range(cfg.getcTNUPADFIN & Target.row).column) Then
                utils.setTNUNNC Range(cfg.getcTNUNNC & Target.row).Value, oldTnunnc, Target.row
            End If
            
            ActiveWorkbook.Sheets(4).Visible = xlSheetVeryHidden
            ActiveWorkbook.Sheets(2).Select
            Application.ScreenUpdating = True
            Application.Cursor = xlDefault
        End If
    End If
End Sub

Sub allowEventHandling()

    globals.setAllowEventHandling (True)
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub
