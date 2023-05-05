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

Private Sub MultiPage_Change()

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
        Range("C8").Value = lstSupplierResults.Value
    End If
End Sub

Private Sub lstSupplierResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(lstSupplierResults.Value) Then
        Range("C8").Value = lstSupplierResults.Value
        frmSearch.Hide
    End If
End Sub

Private Sub btnSearchSuppliers_Click()
    Application.Cursor = xlWait
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.searchSuppliers(txtSupplierCode.Value, txtSupplierName.Value)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    lstSupplierResults.Clear
    If rs.EOF = True Then
        MsgBox "Tražena pretraga nije dala rezulat", vbOKOnly, "Informacija"
    End If
    
    functions.insertLog "search_supplier", _
    "{ supplierCode: " & txtSupplierCode.Value _
    & ", supplierName: " & txtSupplierName.Value _
    & " }" _
    , CStr(SQLStr)
    
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
