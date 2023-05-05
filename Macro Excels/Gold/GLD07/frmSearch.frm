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
            loadNtars
        Application.Cursor = xlDefault
    End If
End Sub

Private Sub lstNtarResults_Click()
    If Not IsNull(lstNtarResults.Value) Then
        Range("C8").Value = lstNtarResults.Value
        Range("C9").ClearContents
        Range("C10").ClearContents
    End If
End Sub
Private Sub lstNtarResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstNtarResults.Value) Then
        Range("C8").Value = lstNtarResults.Value
        Range("C9").ClearContents
        Range("C10").ClearContents
        frmSearch.Hide
    End If
End Sub

Private Sub loadNtars()

    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.selectNtars()
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    lstNtarResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_ntars", "", CStr(SQLStr)
    
    Do Until rs.EOF = True
        lstNtarResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault

End Sub

Private Sub txtSiteCode_Change()
    If Len(txtSiteCode.Value) > 0 Then
        txtSiteName.Value = ""
    End If
End Sub

Private Sub txtSiteName_Change()
    If Len(txtSiteName.Value) > 0 Then
        txtSiteCode.Value = ""
    End If
End Sub
Private Sub lstSiteResults_Click()
    If Not IsNull(lstSiteResults.Value) Then
        Range("C9").Value = lstSiteResults.Value
        Range("C8").ClearContents
        Range("C10").ClearContents
    End If
End Sub
Private Sub lstSiteResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstSiteResults.Value) Then
        Range("C9").Value = lstSiteResults.Value
        Range("C8").ClearContents
        Range("C10").ClearContents
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchSites_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.searchSites(txtSiteCode.Value, txtSiteName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    lstSiteResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_site", _
    "{ siteCode: " & txtSiteCode.Value _
    & ", siteName: " & txtSiteName.Value _
    & " }" _
    , CStr(SQLStr)
    
    Do Until rs.EOF = True
        lstSiteResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
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
        Range("C10").Value = lstArticleResults.Value
        Range("C9").ClearContents
        Range("C8").ClearContents
        Range("C12").ClearContents
    End If
End Sub
Private Sub lstArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleResults.Value) Then
        Range("C10").Value = lstArticleResults.Value
        Range("C9").ClearContents
        Range("C8").ClearContents
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
    
    SQLStr = queries.searchArticles(txtArticleCode.Value, txtArticleName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    lstArticleResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_article", _
    "{ articleCode: " & txtArticleCode.Value _
    & ", articleName: " & txtArticleName.Value _
    & " }" _
    , CStr(SQLStr)
    
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
    End If
End Sub
Private Sub lstMSNodeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstMSNodeResults.Value) Then
        Range("C12").Value = lstMSNodeResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchMSNode_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.searchMSNodes(txtMSCode.Value, txtMSName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    lstMSNodeResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If

    functions.insertLog "search_ms_node", _
    "{ msCode: " & txtMSCode.Value _
    & ", msName: " & txtMSName.Value _
    & " }" _
    , CStr(SQLStr)
    
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
