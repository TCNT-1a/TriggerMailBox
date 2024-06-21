VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1


Private Sub Application_Startup()
    Debug.Print olFolderInbox
    Set Items = Session.GetDefaultFolder(olFolderInbox).Items
End Sub


Private Sub Items_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item
        Dim s As String
        s = LCase(mail.subject)
        Debug.Print mail.subject
        If mail.UnRead And InStr(s, "[rpa]") > 0 Then
            Shell "C:\RPA\run1.bat", vbNormalFocus

'            MsgBox "B?n có email m?i t?: " & mail.SenderName & vbCrLf & "Ch? d?: " & mail.subject
'
'            mail.UnRead = False
'            mail.Save
        End If
    End If
End Sub



Sub InsertDataToSQLServer()
    
    Dim conn As Object
    Dim cmd As Object
    Dim connStr As String
    Dim sqlStr As String
    Dim serverName As String
    Dim databaseName As String
    Dim userId As String
    Dim password As String

    
    serverName = "N\MSSQLSERVER01"
    databaseName = "UiPath"
    userId = "sa" ' Change this to your user ID
    password = "123456" ' Change this to your password
    
    connStr = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & databaseName & ";User ID=" & userId & ";Password=" & password & ";"
    
    ' Create the connection and command objects
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' Open the connection
    conn.Open connStr
    
    ' Set the command text (SQL query)
    'sqlStr = "INSERT INTO UiPathRequest (column1, column2) VALUES ('value1', 'value2')" ' Change this to your SQL query
    '
    sqlStr = "select * from UiPathRequest"
    ' Set the command properties
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sqlStr
    cmd.CommandType = 1 ' adCmdText
    
    ' Execute the command
    cmd.Execute
    
    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    
    Debug.Print "Data inserted successfully!"
End Sub

