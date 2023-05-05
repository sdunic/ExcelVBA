Attribute VB_Name = "db"
Dim connectionString As String

Private Function getServer() As String
    getServer = "ERP-DWH"
End Function

Function getOracleServer() As String
    getOracleServer = "ORACLESERVER"
End Function

Function getDatabase() As String
    getDatabase = "TommyICT"
End Function

Private Function getUser() As String
    getUser = "ExcelApp"
End Function

Private Function getPassword() As String
    getPassword = "sqlexcel"
End Function

Public Function getDocType() As String
    getDocType = "alat"
End Function

Public Function getDocName() As String
    getDocName = "gift_card"
End Function

Public Function getDocVersion() As String
    getDocVersion = "v 1.00"
End Function

Function getConnectionString() As String
    server = getServer
    database = getDatabase
    user = getUser
    pass = getPassword
    getConnectionString = "Driver={SQL Server};Server=" & server & ";Database=" & database & ";Uid=" & user & ";Pwd=" & pass & ";"
End Function

