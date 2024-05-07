Option Explicit

Sub ConnectToAzureSQL()
    Dim conn As Object
    Dim connectionString As String
    
    ' Connection string
    connectionString = "Provider=SQLOLEDB;" & _
                    "Data Source=chatbotserver456.database.windows.net;" & _
                    "Initial Catalog=pocdb;" & _
                    "User Id=sqlserver;" & _
                    "Password=chatbot@123;"

    ' Create a new connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open the connection
    conn.Open connectionString
    
    ' Check if the connection is successful
    If conn.State = 1 Then
        MsgBox "Connected to Azure SQL Database successfully!", vbInformation
    Else
        MsgBox "Failed to connect to Azure SQL Database!", vbExclamation
    End If
    
    ' Close the connection
    conn.Close
    Set conn = Nothing
End Sub