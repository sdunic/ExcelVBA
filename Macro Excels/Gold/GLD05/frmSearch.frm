VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "ODABIR PARAMETARA"
   ClientHeight    =   3810
   ClientLeft      =   210
   ClientTop       =   810
   ClientWidth     =   11655
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    MultiPage_Change
End Sub
Private Sub MultiPage_Change()
    If MultiPage.Value = 0 Then
        Application.Cursor = xlWait
        If lstPriceListTypes.ListCount = 0 Then
            lstPriceListTypes.AddItem "3500" & " - " & "REDOVAN CJENIK"
            lstPriceListTypes.AddItem "3000" & " - " & "AKCIJSKI CJENIK"
            lstPriceListTypes.AddItem "2000" & " - " & "RASPRODAJA"
            lstPriceListTypes.AddItem "1000" & " - " & "ISTEK ROKA"
        End If
        Application.Cursor = xlDefault
    End If
End Sub

'CJENICI I TRGOVINE
Private Sub lstPriceListTypes_Change()
    If lstPriceListTypes.ListCount = 4 Then
        Application.Cursor = xlWait
        Application.ScreenUpdating = False
        Range("C7").Value = lstPriceListTypes.Value
        Range("C8").ClearContents
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
                
        SQLstr = queries.searchStores(CStr(Split(Range("C7").Value, " - ")(0)))
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        lstStoreResults.Clear
        If rs.EOF = True Then
            MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
        End If
        
        functions.insertLog "search_ntars", _
        "{ ntarType: " & lstPriceListTypes.Value _
        & " }" _
        , CStr(SQLstr)
        
        Do Until rs.EOF = True
            lstStoreResults.AddItem rs(1)
            rs.MoveNext
        Loop
                        
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
    End If
End Sub

Private Sub lstStoreResults_Change()
    Range("C8").ClearContents
    For i = 0 To lstStoreResults.ListCount - 1
        If lstStoreResults.Selected(i) = True Then
            Range("C8").Value = Range("C8").Value + Left(lstStoreResults.List(i), 5) + ";"
        End If
    Next i
    If Len(Range("C8").Value) > 0 Then
        Range("C8").Value = Left(Range("C8").Value, Len(Range("C8").Value) - 1)
    End If
End Sub

'DOBAVLJAÈ
Private Sub txtSupplierCode_Change()
    If Len(txtSupplierCode.Value) > 0 Then
        txtSupplierName.Value = ""
    End If
End Sub

Private Sub txtSupplierName_Change()
    If Len(txtSupplierName.Value) > 0 Then
        txtSupplierCode.Value = ""
    End If
End Sub
Private Sub lstSupplierResults_Click()
    If Not IsNull(lstSupplierResults.Value) Then
        Range("C10").Value = lstSupplierResults.Value
    End If
End Sub
Private Sub lstSupplierResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstSupplierResults.Value) Then
        Range("C10").Value = lstSupplierResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchSuppliers_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchSuppliers(txtSupplierCode.Value, txtSupplierName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstSupplierResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If

    functions.insertLog "search_supplier", _
    "{ supplierCode: " & txtSupplierCode.Value _
    & ", supplierName: " & txtSupplierName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstSupplierResults.AddItem rs(0) & " - " & rs(2) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub


'ROBNI ÈVOR
Private Sub txtMSCode_Change()
    If Len(txtMSCode.Value) > 0 Then
        txtMSName.Value = ""
    End If
End Sub
Private Sub txtMSName_Change()
    If Len(txtMSName.Value) > 0 Then
        txtMSCode.Value = ""
    End If
End Sub
Private Sub lstMSNodeResults_Click()
    If Not IsNull(lstMSNodeResults.Value) Then
        Range("C12").Value = lstMSNodeResults.Value
        Range("C13").ClearContents
    End If
End Sub
Private Sub lstMSNodeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstMSNodeResults.Value) Then
        Range("C12").Value = lstMSNodeResults.Value
        Range("C13").ClearContents
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchMSNode_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchMSNodes(txtMSCode.Value, txtMSName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstMSNodeResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
     functions.insertLog "search_supplier", _
    "{ MSCode: " & txtMSCode.Value _
    & ", MSName: " & txtMSName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstMSNodeResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub

'ARTIKL
Private Sub txtArticleCode_Change()
    If Len(txtArticleCode.Value) > 0 Then
        txtArticleName.Value = ""
    End If
End Sub
Private Sub txtArticleName_Change()
    If Len(txtArticleName.Value) > 0 Then
        txtArticleCode.Value = ""
    End If
End Sub
Private Sub lstArticleResults_Click()
    If Not IsNull(lstArticleResults.Value) Then
        Range("C13").Value = lstArticleResults.Value
        Range("C12").ClearContents
    End If
End Sub
Private Sub lstArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleResults.Value) Then
        Range("C13").Value = lstArticleResults.Value
        Range("C12").ClearContents
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchArticle_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchArticles(txtArticleCode.Value, txtArticleName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstArticleResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
        
    functions.insertLog "search_article", _
    "{ articleCode: " & txtArticleCode.Value _
    & ", articleName: " & txtArticleName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstArticleResults.AddItem rs(0) & " - " & rs(1) & " - " & rs(3)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub
