VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public subject As String
Public body As String
Public fileAttachment As Collection
Public Sub init()
Set fileAttachment = New Collection
End Sub
