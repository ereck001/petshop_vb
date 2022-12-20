Attribute VB_Name = "db"
Dim conn As Object
Function AbreConn() As Object

    Set conn = New ADODB.Connection

    conn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=petshop;UID=sa;PWD=<YourStrong@Passw0rd>;"
    conn.Open
    
    Set AbreConn = conn
    
End Function

