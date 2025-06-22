# Outlook VBA Macros for Productivity and Safety

Here's a collection of Outlook VBA macros, including safety confirmation system and many other productivity tools for daily office routines:
>[!Note]
>```The Developer tab isn't available in the new Outlook for Windows```  
>```Use Outlook Classic to enable Developer for macros```   
![image](https://github.com/user-attachments/assets/8ba400cc-82e0-4aa0-82db-7a28a8fd3ca6)


## Safety & Emergency Macros (5)

```vba
' 1. Safety Check-In System (your requested macro)
Sub SafetyCheckInSystem()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Outlook.MailItem
    Dim response As VbMsgBoxResult
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderInbox)
    
    ' Check if 24 hours have passed since last confirmation
    If DateDiff("h", GetLastConfirmationTime(), Now) >= 24 Then
        response = MsgBox("You haven't confirmed your safety in 24 hours. Are you OK?" & vbCrLf & _
                         "Click YES to confirm you're safe." & vbCrLf & _
                         "Click NO to send emergency alerts.", vbYesNo + vbCritical, "Safety Check")
        
        If response = vbNo Then
            ' Send emergency message
            Set olItem = olApp.CreateItem(olMailItem)
            With olItem
                .Subject = "EMERGENCY: I NEED HELP"
                .Body = "Hey team," & vbCrLf & vbCrLf & _
                       "I went to a risk area yesterday. If you're seeing this message, it means I haven't checked in and may need assistance." & vbCrLf & vbCrLf & _
                       "My last known location was: [INSERT LOCATION]" & vbCrLf & _
                       "My emergency contact is: [INSERT CONTACT]" & vbCrLf & vbCrLf & _
                       "Please alert the appropriate authorities." & vbCrLf & vbCrLf & _
                       "Sent automatically by my Outlook safety system."
                .Importance = olImportanceHigh
                
                ' Add emergency contacts (modify as needed)
                .Recipients.Add "security@company.com"
                .Recipients.Add "manager@company.com"
                .Recipients.Add "team@company.com"
                
                .Send
            End With
            MsgBox "Emergency alert sent to your contacts!", vbExclamation, "Alert Sent"
        Else
            ' Update last confirmation time
            UpdateLastConfirmationTime
            MsgBox "Safety confirmation received. Thank you!", vbInformation, "Confirmed"
        End If
    End If
End Sub

Function GetLastConfirmationTime() As Date
    ' Store/retrieve last confirmation time in a custom Outlook property
    On Error Resume Next
    GetLastConfirmationTime = Outlook.Application.Session.DefaultStore.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/LastSafetyConfirm")
    If Err.Number <> 0 Then
        GetLastConfirmationTime = Now ' Default to now if never confirmed
    End If
    On Error GoTo 0
End Function

Sub UpdateLastConfirmationTime()
    ' Store the current time as last confirmation time
    Outlook.Application.Session.DefaultStore.SetProperty "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/LastSafetyConfirm", Now
End Sub
```
```vba
' 2. Send delayed "I'm safe" message
Sub ScheduleSafetyConfirmation()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olItem = olApp.CreateItem(olMailItem)
    
    With olItem
        .Subject = "I'm Safe - Automated Message"
        .Body = "Hello," & vbCrLf & vbCrLf & _
               "This is an automated message to confirm I'm safe after my assignment." & vbCrLf & vbCrLf & _
               "Best regards," & vbCrLf & _
               "[Your Name]"
        
        ' Add recipients
        .Recipients.Add "manager@company.com"
        .Recipients.Add "team@company.com"
        
        ' Set deferred delivery time (24 hours from now)
        .DeferredDeliveryTime = DateAdd("h", 24, Now)
        
        ' Save to drafts (will send automatically at scheduled time)
        .Save
    End With
    
    MsgBox "Safety confirmation message scheduled to send in 24 hours.", vbInformation, "Scheduled"
End Sub
```
```vba
' 3. Emergency contact quick email
Sub EmergencyContact()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olItem = olApp.CreateItem(olMailItem)
    
    With olItem
        .Subject = "URGENT: Emergency Assistance Needed"
        .Body = "Dear Emergency Contact," & vbCrLf & vbCrLf & _
               "I require immediate assistance. Here are my details:" & vbCrLf & vbCrLf & _
               "Current Location: [INSERT LOCATION]" & vbCrLf & _
               "Situation: [DESCRIBE EMERGENCY]" & vbCrLf & vbCrLf & _
               "Please contact me at: [YOUR PHONE NUMBER]" & vbCrLf & vbCrLf & _
               "Sent via Outlook emergency button."
        .Importance = olImportanceHigh
        
        ' Add predefined emergency contacts
        .Recipients.Add "emergency@company.com"
        .Recipients.Add "security@company.com"
        
        .Display ' Show to user before sending
    End With
End Sub
```
```vba
' 4. Check-in reminder system
Sub SetupCheckInReminders()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.TaskItem
    
    Set olApp = Outlook.Application
    
    ' Create recurring task for check-ins
    Set olItem = olApp.CreateItem(olTaskItem)
    With olItem
        .Subject = "SAFETY CHECK-IN REQUIRED"
        .Body = "Remember to check in every 24 hours when working in risk areas."
        .StartDate = Date
        .DueDate = Date
        .ReminderSet = True
        .ReminderTime = DateAdd("h", 20, Now) ' Remind after 20 hours
        .RecurrenceState = olTaskRecurDaily
        .Save
    End With
    
    MsgBox "Daily safety check-in reminders have been set up.", vbInformation, "Reminders Added"
End Sub
```
```vba
' 5. Location tracker email
Sub SendLocationUpdate()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olItem = olApp.CreateItem(olMailItem)
    
    With olItem
        .Subject = "Current Location Update"
        .Body = "Team," & vbCrLf & vbCrLf & _
               "As part of my safety protocol, here's my current location:" & vbCrLf & vbCrLf & _
               "Location: [INSERT LOCATION]" & vbCrLf & _
               "Time: " & Format(Now, "yyyy-mm-dd hh:mm") & vbCrLf & _
               "Expected return: [INSERT TIME]" & vbCrLf & vbCrLf & _
               "This is an automated location update."
        .Importance = olImportanceNormal
        
        ' Add recipients
        .Recipients.Add "team@company.com"
        .Recipients.Add "manager@company.com"
        
        .Display ' Show to user before sending
    End With
End Sub
```

## Email Management Macros (9)

```vba
' 6. Quick category and move
Sub CategorizeAndMoveSelected()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olItem As Object
    Dim olFolder As Outlook.MAPIFolder
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Get selected items
    If olApp.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No items selected", vbExclamation
        Exit Sub
    End If
    
    ' Let user choose category
    Dim category As String
    category = InputBox("Enter category name:", "Categorize Emails")
    If category = "" Then Exit Sub
    
    ' Let user choose destination folder
    Set olFolder = olApp.Session.PickFolder
    If olFolder Is Nothing Then Exit Sub
    
    ' Process each selected item
    For Each olItem In olApp.ActiveExplorer.Selection
        If TypeName(olItem) = "MailItem" Then
            olItem.Categories = category
            olItem.Save
            olItem.Move olFolder
        End If
    Next
    
    MsgBox "Processed " & olApp.ActiveExplorer.Selection.Count & " items", vbInformation
End Sub
```
```vba
' 7. Clean up old emails
Sub AutoCleanOldEmails()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim daysOld As Integer
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderInbox)
    
    daysOld = 30 ' Default to 30 days
    count = 0
    
    ' Process all items in Inbox
    For Each olItem In olFolder.Items
        If TypeName(olItem) = "MailItem" Then
            If DateDiff("d", olItem.ReceivedTime, Now) > daysOld Then
                olItem.Delete
                count = count + 1
            End If
        End If
    Next
    
    MsgBox "Deleted " & count & " emails older than " & daysOld & " days", vbInformation
End Sub
```
```vba
' 8. Quick reply template
Sub QuickReplyTemplate()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    Dim olReply As Outlook.MailItem
    
    Set olApp = Outlook.Application
    
    If Not TypeName(olApp.ActiveWindow) = "Inspector" Then
        MsgBox "Please open an email first", vbExclamation
        Exit Sub
    End If
    
    Set olItem = olApp.ActiveInspector.CurrentItem
    Set olReply = olItem.Reply
    
    With olReply
        .Body = "Thank you for your email." & vbCrLf & vbCrLf & _
               "I've received your message and will respond more fully soon." & vbCrLf & vbCrLf & _
               "Best regards," & vbCrLf & _
               Application.Session.CurrentUser.Name
        .Display
    End With
End Sub
```
```vba
' 9. Send email with attachment from selected folder
Sub SendFilesFromFolder()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    Dim fldr As FileDialog
    Dim file As Variant
    
    Set olApp = Outlook.Application
    Set olItem = olApp.CreateItem(olMailItem)
    
    ' Let user select files
    Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    With fldr
        .Title = "Select files to attach"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
        
        ' Create email
        With olItem
            .Subject = "Files as requested"
            .Body = "Hello," & vbCrLf & vbCrLf & "Please find attached the requested files." & vbCrLf & vbCrLf & "Regards," & vbCrLf & Application.Session.CurrentUser.Name
            
            ' Add attachments
            For Each file In .SelectedItems
                .Attachments.Add file
            Next
            
            .Display
        End With
    End With
End Sub
```
```vba
' 10. Convert selected emails to tasks
Sub EmailsToTasks()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olItem As Object
    Dim olTask As Outlook.TaskItem
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    If olApp.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No emails selected", vbExclamation
        Exit Sub
    End If
    
    For Each olItem In olApp.ActiveExplorer.Selection
        If TypeName(olItem) = "MailItem" Then
            Set olTask = olApp.CreateItem(olTaskItem)
            With olTask
                .Subject = "RE: " & olItem.Subject
                .Body = "Email from " & olItem.SenderName & " on " & olItem.ReceivedTime & vbCrLf & vbCrLf & olItem.Body
                .StartDate = Date
                .DueDate = Date + 7 ' Due in 1 week
                .ReminderSet = True
                .ReminderTime = Date + 1 ' Remind tomorrow
                .Save
            End With
            olItem.UnRead = False
            olItem.Save
        End If
    Next
    
    MsgBox "Created " & olApp.ActiveExplorer.Selection.Count & " tasks", vbInformation
End Sub
```
```vba
' 11. Archive entire conversation
Sub ArchiveConversation()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    Dim olNS As Outlook.NameSpace
    Dim olArchiveFolder As Outlook.MAPIFolder
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Check if an email is selected
    If Not TypeName(olApp.ActiveWindow) = "Inspector" Then
        MsgBox "Please open an email first", vbExclamation
        Exit Sub
    End If
    
    Set olItem = olApp.ActiveInspector.CurrentItem
    
    ' Create or get Archive folder
    On Error Resume Next
    Set olArchiveFolder = olNS.GetDefaultFolder(olFolderInbox).Folders("Archive")
    If olArchiveFolder Is Nothing Then
        Set olArchiveFolder = olNS.GetDefaultFolder(olFolderInbox).Folders.Add("Archive")
    End If
    On Error GoTo 0
    
    ' Find and move all items in the conversation
    Dim convItems As Outlook.Items
    Set convItems = olItem.GetConversation.GetRootItems
    
    Dim convItem As Object
    For Each convItem In convItems
        If TypeName(convItem) = "MailItem" Then
            convItem.Move olArchiveFolder
        End If
    Next
    
    MsgBox "Archived " & convItems.Count & " emails from this conversation", vbInformation
End Sub
```
```vba
' 12. Follow up in X days
Sub FollowUpInDays()
    Dim olApp As Outlook.Application
    Dim olItem As Outlook.MailItem
    Dim days As String
    
    Set olApp = Outlook.Application
    
    ' Check if an email is selected
    If Not TypeName(olApp.ActiveWindow) = "Inspector" Then
        MsgBox "Please open an email first", vbExclamation
        Exit Sub
    End If
    
    Set olItem = olApp.ActiveInspector.CurrentItem
    
    days = InputBox("Follow up in how many days?", "Follow Up", "7")
    If days = "" Then Exit Sub
    If Not IsNumeric(days) Then
        MsgBox "Please enter a number", vbExclamation
        Exit Sub
    End If
    
    ' Create follow up task
    Dim olTask As Outlook.TaskItem
    Set olTask = olApp.CreateItem(olTaskItem)
    With olTask
        .Subject = "Follow up: " & olItem.Subject
        .Body = "Follow up on email from " & olItem.SenderName & vbCrLf & vbCrLf & olItem.Body
        .StartDate = Date + CInt(days) - 1
        .DueDate = Date + CInt(days)
        .ReminderSet = True
        .ReminderTime = Date + CInt(days)
        .Categories = "Follow Up"
        .Save
    End With
    
    MsgBox "Follow up task created for " & Format(Date + CInt(days), "mmmm d, yyyy"), vbInformation
End Sub
```
```vba
' 13. Extract all attachments from selected emails
Sub ExtractAllAttachments()
    Dim olApp As Outlook.Application
    Dim olItem As Object
    Dim olAtt As Outlook.Attachment
    Dim savePath As String
    Dim count As Integer
    
    Set olApp = Outlook.Application
    
    If olApp.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No emails selected", vbExclamation
        Exit Sub
    End If
    
    ' Get save folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to save attachments"
        If .Show <> -1 Then Exit Sub
        savePath = .SelectedItems(1) & "\"
    End With
    
    count = 0
    
    ' Process each selected email
    For Each olItem In olApp.ActiveExplorer.Selection
        If TypeName(olItem) = "MailItem" Then
            For Each olAtt In olItem.Attachments
                olAtt.SaveAsFile savePath & olAtt.FileName
                count = count + 1
            Next
        End If
    Next
    
    MsgBox "Saved " & count & " attachments to " & savePath, vbInformation
End Sub
```
