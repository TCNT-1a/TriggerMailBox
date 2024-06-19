Attribute VB_Name = "Module1"
Sub main()
    Dim root As Outlook.Folder
    Set root = olRootFolder
    Dim rpa As Outlook.Folder
    Set rpa = findFolder(root, "RPA")
    ListEmailsInFolder rpa
    
    
    Set root = Nothing
    Set rpa = Nothing
End Sub

Function olRootFolder()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim i As Integer
    
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olRootFolder = olNamespace.GetDefaultFolder(olFolderInbox).Parent
    
    Set olNamespace = Nothing
    Set olApp = Nothing
End Function

Function findFolder(olFolder As Outlook.Folder, folderName As String) As Outlook.Folder
    
    If olFolder.Name = folderName Then
        Set findFolder = olFolder
        Exit Function
    End If
    
    Dim subFolder As Outlook.Folder
    For Each subFolder In olFolder.Folders
        Set findFolder = findFolder(subFolder, folderName)
        If Not findFolder Is Nothing Then Exit Function
    Next subFolder
    
    Set findFolder = Nothing
End Function


Sub ListEmailsInFolder(olFolder As Outlook.Folder)
    Dim olMail As Outlook.MailItem
    For i = 1 To olFolder.Items.Count
        Set email = olFolder.Items(i)
        If TypeOf email Is Outlook.MailItem Then
            If email.UnRead Then
                Set olMail = email
                Debug.Print "Subject: " & olMail.subject
                Debug.Print "Received: " & olMail.ReceivedTime
                Debug.Print "Sender: " & olMail.SenderName
                olMail.MessageClas
                For Each Attachment In olMail.Attachments
            
                    ' Construct the file path for saving the attachment
                    FilePath = "C:\RPA_Temp\" & Attachment.FileName
                    Attachment.SaveAsFile FilePath
                    Debug.Print "Attachment saved to: " & FilePath
                Next Attachment
                
            End If
            
            Debug.Print "---------------------------------"
        End If
    Next i
    ' Clean up
    Set olMail = Nothing
    Set olFolder = Nothing

End Sub

