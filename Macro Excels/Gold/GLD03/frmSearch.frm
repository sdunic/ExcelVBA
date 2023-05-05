VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "ODABIR PARAMETARA"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11715
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
        Range("C7").Value = lstMSNodeResults.Value
        Range("C9").ClearContents
    End If
End Sub
Private Sub lstMSNodeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstMSNodeResults.Value) Then
        Range("C7").Value = lstMSNodeResults.Value
        Range("C9").ClearContents
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchMSNode_Click()
        
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
    
    functions.insertLog "search_ms_node", _
    "{ msCode: " & txtMSCode.Value _
    & ", msName: " & txtMSName.Value _
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
       
End Sub




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
        Range("C7").ClearContents
        Range("C9").Value = lstArticleResults.Value
    End If
End Sub
Private Sub lstArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleResults.Value) Then
        Range("C7").ClearContents
        Range("C9").Value = lstArticleResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchArticles_Click()
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
End Sub




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
        Range("C11").Value = lstSupplierResults.Value
    End If
End Sub
Private Sub lstSupplierResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstSupplierResults.Value) Then
        Range("C11").Value = lstSupplierResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchSuppliers_Click()
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
End Sub
