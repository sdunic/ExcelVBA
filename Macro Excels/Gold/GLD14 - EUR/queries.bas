Attribute VB_Name = "queries"
Function getLog(doc_type As String, doc_name As String, doc_version As String, domain_user As String, operation As String, parameters As String, query As String) As String
    getLog = "INSERT INTO [excel].[excel_logovi] (vrsta, naziv, verzija, korisnik, operacija, parametri, datum_vrijeme, sql_upit) VALUES " _
            & "('" & doc_type & "', '" & doc_name & "', '" & doc_version & "', '" & domain_user & "', '" & operation & ""
    
    If Len(parameters) > 0 Then
        getLog = getLog & "', '" & parameters & "',"
    Else
        getLog = getLog & "', " & "NULL" & ","
    End If
    
    getLog = getLog & " current_timestamp"
    
    If Len(query) > 0 Then
        getLog = getLog & ", '" & query & "'"
    Else
        getLog = getLog & ", " & "NULL"
    End If
    
    getLog = getLog & ")"
    
End Function

Function getDocumentVersion(doc_name As String) As String

     getDocumentVersion = "SELECT TOP 1 [document_version] FROM [excel].[excel_document_versions] WHERE [document_name] = '" & doc_name & "'"
     getDocumentVersion = getDocumentVersion + " ORDER BY [timestamp] DESC"
     
End Function

Function GetMPCdata() As String
    GetMPCdata = "EXEC [" + db.getDatabase + "].[" + db.getProcedurePrefix + "].[" + db.getProcedure + "]"
End Function
