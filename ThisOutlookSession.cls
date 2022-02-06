VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    Sub myRuleMacro(item As Outlook.MailItem)
    Dim objDestFolder As Folder
    Dim objFolder As Folder
    Dim objDictionary As Object
    Dim i As Long
    Dim objItem As Object
    Dim strKey As String

    Set objDictionary = CreateObject("scripting.dictionary")
    'Select a source folder
    Set objFolder = Application.GetNamespace("MAPI").Folders("EMAIL-ADDDRESS-HERE@EXAMPLE.CA").Folders("Inbox").Folders("Printer Status Alerts")
    'Select a destination folder
    Set objDestFolder = Application.GetNamespace("MAPI").Folders("EMAIL-ADDDRESS-HERE@EXAMPLE.CA").Folders("Inbox").Folders("Printer Status Alerts").Folders("Duplicate Alerts")

    If Not (objFolder Is Nothing) Then
       For i = objFolder.Items.Count To 1 Step -1
           Set objItem = objFolder.Items.item(i)
 
           Select Case objFolder.DefaultItemType
                  'Check email subject, body and sent time
                  Case olMailItem
                       strKey = objItem.Subject & ","
           End Select
 
           strKey = Replace(strKey, ", ", Chr(32))
 
           'Move the duplicate items
           If objDictionary.Exists(strKey) = True Then
              objItem.Move objDestFolder
           Else
              objDictionary.Add strKey, True
           End If
       Next i
    End If
End Sub
