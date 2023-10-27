VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchRows 
   Caption         =   "STAVKE FAKTURE"
   ClientHeight    =   3810
   ClientLeft      =   270
   ClientTop       =   990
   ClientWidth     =   11595
   OleObjectBlob   =   "frmSearchRows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    cfg.Init
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
        Range(cfg.get_artikl & ActiveCell.row).Value = lstArticleResults.Value
        Range(cfg.get_lv_lu & ActiveCell.row).Value = Split(lstArticleResults.Value, " | ")(3) & " | SKU"
    End If
End Sub
Private Sub lstArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleResults.Value) Then
        Range(cfg.get_artikl & ActiveCell.row).Value = lstArticleResults.Value
        Range(cfg.get_lv_lu & ActiveCell.row).Value = Split(lstArticleResults.Value, " | ")(3) & " | SKU"
        frmSearchRows.Hide
    End If
End Sub
Private Sub btnSearchArticle_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    sqlstr = queries.searchArticles(txtArticleCode.Value, txtArticleName.Value)
    'Debug.Print (Sqlstr)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstArticleResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_rows_articles", _
    "{ articleCode: " & txtArticleCode.Value _
    & ", articleName: " & txtArticleName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstArticleResults.AddItem rs(0) & " | " & rs(1) & " | " & rs(3) & " | " & rs(2)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub



Private Sub txtLocationCode_Change()
    If Len(txtLocationCode.Value) > 0 Then
        txtLocationName.Value = ""
    End If
End Sub

Private Sub txtLocationName_Change()
    If Len(txtLocationName.Value) > 0 Then
        txtLocationCode.Value = ""
    End If
End Sub
Private Sub lstLocationResults_Click()
    If Not IsNull(lstLocationResults.Value) Then
        Range(cfg.get_tm & ActiveCell.row).Value = Split(lstLocationResults.Value, " | ")(0)
    End If
End Sub
Private Sub lstLocationResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstLocationResults.Value) Then
        Range(cfg.get_tm & ActiveCell.row).Value = Split(lstLocationResults.Value, " | ")(0)
        frmSearchRows.Hide
    End If
End Sub
Private Sub btnSearchLocations_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    sqlstr = queries.searchLocations(txtLocationCode.Value, txtLocationName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstLocationResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If

    functions.insertLog "search_rows_tm_pm", _
    "{ locationCode: " & txtLocationCode.Value _
    & ", locationName: " & txtLocationName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstLocationResults.AddItem rs(0) & " | " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub


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
    If Not IsNull(lstMSNodeResults.Value) And ActiveCell.row >= cfg.get_stavke Then
        Range(cfg.get_robniCvor & ActiveCell.row).Value = lstMSNodeResults.Value
        Range(cfg.get_analitickiArtikl & ActiveCell.row).ClearContents
    End If
End Sub
Private Sub lstMSNodeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstMSNodeResults.Value) And ActiveCell.row >= cfg.get_stavke Then
        Range(cfg.get_robniCvor & ActiveCell.row).Value = lstMSNodeResults.Value
        Range(cfg.get_analitickiArtikl & ActiveCell.row).ClearContents
        frmSearchRows.Hide
    End If
End Sub
Private Sub btnSearchMSNode_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    sqlstr = queries.searchMSNodes(txtMSCode.Value, txtMSName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstMSNodeResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_rows_analytical_node", _
    "{ MSCode: " & txtMSCode.Value _
    & ", MSName: " & txtMSName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstMSNodeResults.AddItem rs(0) & " | " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub

Private Sub txtAnalyticalArticleCode_Change()
    If Len(txtAnalyticalArticleCode.Value) > 0 Then
        txtAnalyticalArticleName.Value = ""
    End If
End Sub
Private Sub txtAnalyticalArticleName_Change()
    If Len(txtAnalyticalArticleName.Value) > 0 Then
        txtAnalyticalArticleCode.Value = ""
    End If
End Sub
Private Sub lstAnalyticalArticleResults_Click()
    If Not IsNull(lstAnalyticalArticleResults.Value) Then
         Range(cfg.get_robniCvor & ActiveCell.row).ClearContents
        Range(cfg.get_analitickiArtikl & ActiveCell.row).Value = lstAnalyticalArticleResults.Value
    End If
End Sub
Private Sub lstAnalyticalArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstAnalyticalArticleResults.Value) Then
        Range(cfg.get_robniCvor & ActiveCell.row).ClearContents
        Range(cfg.get_analitickiArtikl & ActiveCell.row).Value = lstAnalyticalArticleResults.Value
        frmSearchRows.Hide
    End If
End Sub
Private Sub btnSearchAnalyticalArticle_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    sqlstr = queries.searchAnalyticalArticles(txtAnalyticalArticleCode.Value, txtAnalyticalArticleName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstAnalyticalArticleResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_rows_analytical_article", _
    "{ articleCode: " & txtAnalyticalArticleCode.Value _
    & ", articleName: " & txtAnalyticalArticleName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstAnalyticalArticleResults.AddItem rs(0) & " | " & rs(1) & " | " & rs(3)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub
