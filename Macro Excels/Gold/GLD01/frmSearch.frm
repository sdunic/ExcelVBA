VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "ODABIR PARAMETARA"
   ClientHeight    =   3810
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   11655
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
        Range("C15").Value = lstMSNodeResults.Value
    End If
End Sub
Private Sub lstMSNodeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstMSNodeResults.Value) Then
        Range("C15").Value = lstMSNodeResults.Value
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
        Range("C9").Value = lstLocationResults.Value
    End If
End Sub
Private Sub lstLocationResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstLocationResults.Value) Then
        Range("C9").Value = lstLocationResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchLocations_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchLocations(txtLocationCode.Value, txtLocationName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstLocationResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_location", _
    "{ locationCode: " & txtLocationCode.Value _
    & ", locationName: " & txtLocationName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstLocationResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
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
        chkSelectedSupplier.Caption = lstSupplierResults.Value
        chkSelectedSupplier.Visible = True
    End If
End Sub
Private Sub lstSupplierResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstSupplierResults.Value) Then
        Range("C11").Value = lstSupplierResults.Value
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




Private Sub txtContractCode_Change()
    If Len(txtContractCode.Value) > 0 Then
        txtContractName.Value = ""
    End If
End Sub

Private Sub txtContractName_Change()
    If Len(txtContractName.Value) > 0 Then
        txtContractCode.Value = ""
    End If
End Sub
Private Sub lstContractResults_Click()
    If Not IsNull(lstContractResults.Value) Then
        Range("C13").Value = lstContractResults.Value
    End If
End Sub
Private Sub lstContractResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstContractResults.Value) Then
        Range("C13").Value = lstContractResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub chkSelectedSupplier_Click()
    btnSearchContract_Click
End Sub
Private Sub btnSearchContract_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchContracts(txtContractCode.Value, txtContractName.Value, chkSelectedSupplier.Caption, chkSelectedSupplier)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstContractResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_contracts", _
    "{ contractCode: " & txtContractCode.Value _
    & ", contractName: " & txtContractName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstContractResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub




Private Sub txtArticleListCode_Change()
    If Len(txtArticleListCode.Value) > 0 Then
        txtArticleListName.Value = ""
    End If
End Sub
Private Sub txtArticleListName_Change()
    If Len(txtArticleListName.Value) > 0 Then
        txtArticleListCode.Value = ""
    End If
End Sub
Private Sub lstArticleListResults_Click()
    If Not IsNull(lstArticleListResults.Value) Then
        Range("C17").Value = lstArticleListResults.Value
    End If
End Sub
Private Sub lstArticleListResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleListResults.Value) Then
        Range("C17").Value = lstArticleListResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchArticleList_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchArticleLists(txtArticleListCode.Value, txtArticleListName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstArticleListResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_article_list", _
    "{ articleListCode: " & txtArticleListCode.Value _
    & ", articleListName: " & txtArticleListName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstArticleListResults.AddItem rs(0) & " - " & rs(1)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub





Private Sub txtArticleGroupCode_Change()
    If Len(txtArticleGroupCode.Value) > 0 Then
        txtArticleGroupName.Value = ""
    End If
End Sub
Private Sub txtArticleGroupName_Change()
    If Len(txtArticleGroupName.Value) > 0 Then
        txtArticleGroupCode.Value = ""
    End If
End Sub
Private Sub lstArticleGroupResults_Click()
    If Not IsNull(lstArticleGroupResults.Value) Then
        Range("C18").Value = lstArticleGroupResults.Value
    End If
End Sub
Private Sub lstArticleGroupResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleGroupResults.Value) Then
        Range("C18").Value = lstArticleGroupResults.Value
        frmSearch.Hide
    End If
End Sub
Private Sub btnSearchArticleGroup_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.searchArticleGroups(txtArticleGroupCode.Value, txtArticleGroupName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
    lstArticleGroupResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_article_group", _
    "{ articleGroupCode: " & txtArticleGroupCode.Value _
    & ", articleGroupName: " & txtArticleGroupName.Value _
    & " }" _
    , CStr(SQLstr)
    
    Do Until rs.EOF = True
        lstArticleGroupResults.AddItem rs(0) & " - " & rs(1)
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
        Range("C19").Value = lstArticleResults.Value
    End If
End Sub
Private Sub lstArticleResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstArticleResults.Value) Then
        Range("C19").Value = lstArticleResults.Value
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



Private Sub UserForm_Initialize()
    MultiPage_Change
End Sub
Private Sub MultiPage_Change()
    'page 7 je klasa
    If MultiPage.Value = 7 Then
        Application.Cursor = xlWait
        If cbClasses.ListCount = 0 Then
            
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
            
            SQLstr = queries.selectClasses
            
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open SQLstr, Cn, adOpenStatic
            
            cbClasses.Clear
            If rs.EOF = True Then
                MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
            End If
            
            Do Until rs.EOF = True
                cbClasses.AddItem rs(0) & " - " & rs(1)
                rs.MoveNext
            Loop
            
            rs.Close
            Set rs = Nothing
            Cn.Close
            Set Cn = Nothing
        
        End If
        Application.Cursor = xlDefault
    End If
End Sub
Private Sub cbClasses_Change()
        Application.Cursor = xlWait
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.searchClassAttributes(CStr(Split(cbClasses.Value, " - ")(0)))
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        lstAttributeResults.Clear
        If rs.EOF = True Then
            MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
        End If
        
        functions.insertLog "search_class_attributes", _
        "{ class: " & cbClasses.Value _
        & " }" _
        , CStr(SQLstr)
        
        Do Until rs.EOF = True
            lstAttributeResults.AddItem rs(0) & " - " & rs(1)
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
        Application.Cursor = xlDefault
End Sub
Private Sub lstAttributeResults_Click()
    If Not IsNull(lstAttributeResults.Value) Then
        Range("C21").Value = cbClasses.Value
        Range("C22").Value = lstAttributeResults.Value
    End If
End Sub
Private Sub lstAttributeResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstAttributeResults.Value) Then
        Range("C21").Value = cbClasses.Value
        Range("C22").Value = lstAttributeResults.Value
        frmSearch.Hide
    End If
End Sub
