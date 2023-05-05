Attribute VB_Name = "db"
Dim connectionString As String

Function getSSHIP() As String
    'getSSHIP = "172.20.33.60" 'stage
    getSSHIP = "172.20.33.53" 'prod
End Function

Function getSSHport() As Integer
    getSSHport = 22
End Function

Function getSSHuser() As String
    getSSHuser = "egold"
End Function

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

Function getProcedurePrefix() As String
    getProcedurePrefix = "excel"
End Function

Function getProcedure() As String
    getProcedure = "GetGoldRasterOrders_prod"
    'getProcedure = "GetGoldRasterOrders_stage"
End Function

Function getConnectionString() As String
    server = getServer
    database = getDatabase
    user = getUser
    pass = getPassword
    getConnectionString = "Driver={SQL Server};Server=" & server & ";Database=" & database & ";Uid=" & user & ";Pwd=" & pass & ";"
End Function

Public Function getDocType() As String
    getDocType = "alat"
End Function

Public Function getDocName() As String
    getDocName = "raster_orders_prod"
End Function

Public Function getDocVersion() As String
    getDocVersion = "1.11"
End Function


