VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchHeader 
   Caption         =   "ZAGLAVLJE FAKTURE"
   ClientHeight    =   3660
   ClientLeft      =   225
   ClientTop       =   810
   ClientWidth     =   11595
   OleObjectBlob   =   "frmSearchHeader.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    cfg.Init
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
        Range(cfg.get_lokacija & cfg.get_zaglavlje).Value = lstLocationResults.Value
    End If
End Sub
Private Sub lstLocationResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstLocationResults.Value) Then
        Range(cfg.get_lokacija & cfg.get_zaglavlje).Value = lstLocationResults.Value
        frmSearchHeader.Hide
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
    
    functions.insertLog "search_header_location", _
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

Private Sub txtCustomerCode_Change()
    If Len(txtCustomerCode.Value) > 0 Then
        txtCustomerName.Value = ""
    End If
End Sub

Private Sub txtCustomerName_Change()
    If Len(txtCustomerName.Value) > 0 Then
        txtCustomerCode.Value = ""
    End If
End Sub
Private Sub lstCustomerResults_Click()
    If Not IsNull(lstCustomerResults.Value) Then
        Range(cfg.get_kupac & cfg.get_zaglavlje).Value = lstCustomerResults.Value
        chkSelectedCustomer.Visible = True
        chkSelectedCustomer.Caption = lstCustomerResults.Value
        chkSelectedCustomer.Value = True
    End If
End Sub
Private Sub lstCustomerResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstCustomerResults.Value) Then
        Range(cfg.get_kupac & cfg.get_zaglavlje).Value = lstCustomerResults.Value
        chkSelectedCustomer.Visible = True
        chkSelectedCustomer.Caption = lstCustomerResults.Value
        chkSelectedCustomer.Value = True
        frmSearchHeader.Hide
    End If
End Sub
Private Sub btnSearchCustomers_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    sqlstr = queries.searchCustomers(txtCustomerCode.Value, txtCustomerName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstCustomerResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_header_customers", _
    "{ customerCode: " & txtCustomerCode.Value _
    & ", customerName: " & txtCustomerName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstCustomerResults.AddItem rs(0) & " | " & rs(1)
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
        Range(cfg.get_ugovor & cfg.get_zaglavlje).Value = lstContractResults.Value
    End If
End Sub
Private Sub lstContractResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstContractResults.Value) Then
        Range(cfg.get_ugovor & cfg.get_zaglavlje).Value = lstContractResults.Value
        frmSearchHeader.Hide
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
    
    sqlstr = queries.searchContracts(txtContractCode.Value, txtContractName.Value, chkSelectedCustomer.Caption, chkSelectedCustomer)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlstr, Cn, adOpenStatic
    
    lstContractResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_header_contracts", _
    "{ contractCode: " & txtContractCode.Value _
    & ", contractName: " & txtContractName.Value _
    & " }" _
    , CStr(sqlstr)
    
    Do Until rs.EOF = True
        lstContractResults.AddItem rs(0) & " | " & rs(2)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    Application.Cursor = xlDefault
End Sub
