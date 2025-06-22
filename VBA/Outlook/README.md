# Outlook VBA Automation Scripts

<img src="https://upload.wikimedia.org/wikipedia/commons/d/df/Microsoft_Office_Outlook_%282018–present%29.svg" width="15%" height="15%">   
This repository contains a collection of VBA scripts to automate various tasks in Microsoft Outlook, improving productivity and workflow efficiency.   

---
>[!Note]
>```The Developer tab isn't available in the new Outlook for Windows (Outlook Office 365 for Desktop)```  
>```Use Outlook Classic to enable Developer for macros```   
><img src="https://github.com/user-attachments/assets/8ba400cc-82e0-4aa0-82db-7a28a8fd3ca6"  width="50%" height="50%">   

# Outlook VBA Automation Scripts

## Table of Contents
1. [Features](#features)
2. [Script Categories](#script-categories)
   - [Email Automation](#email-automation)
   - [Calendar Management](#calendar-management)
   - [Contact Management](#contact-management)
   - [Task Automation](#task-automation)
   - [Utility Functions](#utility-functions)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Contributing](#contributing)
6. [License](#license)
7. [Support](#support)
8. [Macros](#macros)

## Features
- Automate repetitive Outlook tasks
- Enhance email management workflows
- Streamline calendar operations
- Improve contact organization
- Customizable solutions for various business needs  
--- 
🚧 File Links is Under Construction - Links may not work - Come back in June 30, 2025  
## Script Categories  

### Email Automation  
| Script Name | Description | File Link |
|------------|-------------|----------|
| Auto-Responder | Automatic replies based on specific criteria | [AutoResponder.bas](AutoResponder.bas) |
| Email Categorizer | Automatically categorizes incoming emails | [EmailCategorizer.bas](EmailCategorizer.bas) |
| Bulk Email Processor | Processes multiple emails simultaneously | [BulkEmailProcessor.bas](BulkEmailProcessor.bas) |

### Calendar Management
| Script Name | Description | File Link |
|------------|-------------|----------|
| Meeting Scheduler | Automates meeting scheduling | [MeetingScheduler.bas](MeetingScheduler.bas) |
| Calendar Cleanup | Removes old or duplicate calendar items | [CalendarCleanup.bas](CalendarCleanup.bas) |

### Contact Management
| Script Name | Description | File Link |
|------------|-------------|----------|
| Contact Sync | Synchronizes contacts with external sources | [ContactSync.bas](ContactSync.bas) |
| Group Manager | Manages contact groups efficiently | [GroupManager.bas](GroupManager.bas) |

### Task Automation
| Script Name | Description | File Link |
|------------|-------------|----------|
| Task Reminder | Enhanced reminder system for tasks | [TaskReminder.bas](TaskReminder.bas) |
| Priority Sorter | Automatically prioritizes tasks | [PrioritySorter.bas](PrioritySorter.bas) |

### Utility Functions
| Script Name | Description | File Link |
|------------|-------------|----------|
| Signature Manager | Manages email signatures | [SignatureManager.bas](SignatureManager.bas) |
| Backup Tool | Creates backups of Outlook data | [BackupTool.bas](BackupTool.bas) |

## Installation

1. **Enable Developer Tab**:
   - File → Options → Customize Ribbon
   - Check "Developer" in the right column
   - Click OK

2. **Access VBA Editor**:
   - Press `Alt+F11` or click Developer → Visual Basic

3. **Import Scripts**:
   - In VBA Editor: File → Import File
   - Select the .bas files from this repository

4. **Enable Macros**:
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Enable all macros (for development)

## Usage

1. **Running Scripts**:
   - Most scripts can be run directly from the VBA editor (F5)
   - Some include custom ribbon buttons (see script-specific instructions)

2. **Customization**:
   - Open the script in VBA Editor
   - Modify constants at the top of each script as needed
   - Save changes

3. **Scheduled Automation**:
   - Use Outlook's built-in VBA event handlers (Application_Startup, etc.)
   - Or set up Windows Task Scheduler to run macros at specific times

## Contributing

We welcome contributions! Please follow these guidelines:

1. Fork the repository
2. Create a new branch for your feature (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

Please ensure your code follows existing style conventions and includes appropriate comments.

## License

This project is licensed under the MIT License - see the [LICENSE.md](https://github.com/ThiagoMaria-SecurityIT/Microsoft_Office_Automation_VBA/blob/main/LICENSE) file for details.

## Support

For questions or issues:
- Open an issue on GitHub
- Email: [your email]
- LinkedIn: [your LinkedIn profile]

---

*Last Updated: June 22, 2025*  
*Tested with Outlook (Classic) 2021*

## Macros

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

# Outlook VBA Macros for Security, Compliance, and Audits

Here are VBA macros focused on security information management, compliance tracking, and audit preparation in Outlook:

## 1. Security Awareness Email Reminder
```vba
Sub SendSecurityAwarenessEmail()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .Subject = "Monthly Security Awareness Reminder"
        .HTMLBody = "<p>Hello Team,</p>" & _
                   "<p>As part of our ongoing security awareness program, please remember:</p>" & _
                   "<ul>" & _
                   "<li>Never share your passwords</li>" & _
                   "<li>Verify email senders before clicking links</li>" & _
                   "<li>Lock your workstation when away</li>" & _
                   "<li>Report suspicious emails to security@company.com</li>" & _
                   "</ul>" & _
                   "<p>Thank you for helping keep our organization secure!</p>" & _
                   "<p>Security Team</p>"
        .To = "all-employees@company.com"
        .Importance = olImportanceHigh
        .Send
    End With
    
    MsgBox "Security awareness email sent to all employees.", vbInformation
End Sub
```

## 2. Data Classification Labeling
```vba
Sub ApplyDataClassification()
    Dim olApp As Outlook.Application
    Dim olItem As Object
    Set olApp = Outlook.Application
    
    ' Check if an email is selected
    If TypeName(olApp.ActiveWindow) = "Inspector" Then
        Set olItem = olApp.ActiveInspector.CurrentItem
        
        Dim classification As String
        classification = InputBox("Select classification:" & vbCrLf & _
                                "1. Public" & vbCrLf & _
                                "2. Internal" & vbCrLf & _
                                "3. Confidential" & vbCrLf & _
                                "4. Highly Confidential", "Data Classification", "2")
        
        Select Case classification
            Case "1"
                olItem.Sensitivity = olNormal
                olItem.Subject = "[PUBLIC] " & olItem.Subject
            Case "2"
                olItem.Sensitivity = olNormal
                olItem.Subject = "[INTERNAL] " & olItem.Subject
            Case "3"
                olItem.Sensitivity = olConfidential
                olItem.Subject = "[CONFIDENTIAL] " & olItem.Subject
            Case "4"
                olItem.Sensitivity = olConfidential
                olItem.Subject = "[HIGHLY CONFIDENTIAL] " & olItem.Subject
                olItem.FlagRequest = "Follow up"
                olItem.FlagDueBy = Date + 7
        End Select
        
        olItem.Save
        MsgBox "Classification applied successfully.", vbInformation
    Else
        MsgBox "Please open an email first.", vbExclamation
    End If
End Sub
```

## 3. Compliance Attachment Checker
```vba
Sub CheckForRestrictedAttachments()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim olAtt As Outlook.Attachment
    Dim restrictedTypes As Variant
    Dim foundRestricted As Boolean
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderInbox)
    
    ' Define restricted file types
    restrictedTypes = Array(".exe", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".jar", ".dll")
    
    ' Search last 100 emails
    For Each olItem In olFolder.Items.Restrict("[ReceivedTime] > '" & Format(Date - 7, "ddddd h:nn AMPM") & "'")
        If TypeName(olItem) = "MailItem" Then
            For Each olAtt In olItem.Attachments
                For Each ext In restrictedTypes
                    If LCase(Right(olAtt.FileName, Len(ext))) = ext Then
                        foundRestricted = True
                        MsgBox "Restricted attachment found:" & vbCrLf & _
                               "From: " & olItem.SenderName & vbCrLf & _
                               "Subject: " & olItem.Subject & vbCrLf & _
                               "Attachment: " & olAtt.FileName, vbExclamation, "Security Alert"
                        Exit For
                    End If
                Next
            Next
        End If
    Next
    
    If Not foundRestricted Then
        MsgBox "No restricted attachments found in the last 7 days.", vbInformation
    End If
End Sub
```

## 4. Security Incident Reporting
```vba
Sub ReportSecurityIncident()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim incidentType As String
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    incidentType = InputBox("Select incident type:" & vbCrLf & _
                          "1. Phishing Email" & vbCrLf & _
                          "2. Lost/Stolen Device" & vbCrLf & _
                          "3. Unauthorized Access" & vbCrLf & _
                          "4. Data Leakage" & vbCrLf & _
                          "5. Other", "Security Incident Report")
    
    With olMail
        .To = "security-incident@company.com"
        .Subject = "Security Incident Report - " & Format(Now, "yyyy-mm-dd hh:mm")
        .Body = "Incident Type: " & incidentType & vbCrLf & vbCrLf & _
               "Date/Time: " & Now & vbCrLf & vbCrLf & _
               "Description: " & InputBox("Please describe the incident:", "Incident Details") & vbCrLf & vbCrLf & _
               "Impact: " & InputBox("What is the potential impact?", "Impact Assessment") & vbCrLf & vbCrLf & _
               "Reporter: " & olApp.Session.CurrentUser.Name
        .Importance = olImportanceHigh
        .Display
    End With
End Sub
```

## 5. Backup Important Emails
```vba
Sub BackupSecurityEmails()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim backupFolder As String
    Dim fso As Object
    Dim backupFile As Object
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderInbox)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get backup folder
    backupFolder = Environ("USERPROFILE") & "\Documents\EmailBackup\" & Format(Now, "yyyy-mm-dd") & "\"
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If
    
    count = 0
    
    ' Backup emails with security keywords
    For Each olItem In olFolder.Items.Restrict("[ReceivedTime] > '" & Format(Date - 30, "ddddd h:nn AMPM") & "'")
        If TypeName(olItem) = "MailItem" Then
            If InStr(1, olItem.Subject, "security", vbTextCompare) > 0 Or _
               InStr(1, olItem.Body, "security", vbTextCompare) > 0 Or _
               olItem.Sensitivity = olConfidential Then
                
                Set backupFile = fso.CreateTextFile(backupFolder & "Email_" & Format(Now, "yyyymmddhhnnss") & ".txt")
                backupFile.WriteLine "From: " & olItem.SenderName
                backupFile.WriteLine "Sent: " & olItem.SentOn
                backupFile.WriteLine "To: " & olItem.To
                backupFile.WriteLine "Subject: " & olItem.Subject
                backupFile.WriteLine "Body: " & vbCrLf & olItem.Body
                backupFile.Close
                count = count + 1
            End If
        End If
    Next
    
    MsgBox "Backup completed. " & count & " security-related emails saved to:" & vbCrLf & backupFolder, vbInformation
End Sub
```

## 6. Password Change Reminder
```vba
Sub PasswordChangeReminder()
    Dim olApp As Outlook.Application
    Dim olTask As Outlook.TaskItem
    
    Set olApp = Outlook.Application
    Set olTask = olApp.CreateItem(olTaskItem)
    
    With olTask
        .Subject = "CHANGE YOUR PASSWORD - Quarterly Requirement"
        .Body = "As part of our security policy, you are required to change your password every 90 days." & vbCrLf & _
               "Please update your password for the following systems:" & vbCrLf & _
               "- Network login" & vbCrLf & _
               "- Email account" & vbCrLf & _
               "- Any other company systems" & vbCrLf & vbCrLf & _
               "Remember:" & vbCrLf & _
               "- Use strong passwords (min 12 characters)" & vbCrLf & _
               "- Don't reuse old passwords" & vbCrLf & _
               "- Never share your password"
        .StartDate = Date
        .DueDate = Date + 7
        .ReminderSet = True
        .ReminderTime = Date + 1
        .Categories = "Security"
        .Importance = olImportanceHigh
        .Save
    End With
    
    MsgBox "Password change reminder task created.", vbInformation
End Sub
```

## 7. Audit Trail Generator
```vba
Sub GenerateEmailAuditTrail()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim fso As Object
    Dim auditFile As Object
    Dim startDate As Date
    Dim endDate As Date
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderSentMail)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get date range
    startDate = InputBox("Enter start date (mm/dd/yyyy):", "Audit Period", Format(Date - 30, "mm/dd/yyyy"))
    endDate = InputBox("Enter end date (mm/dd/yyyy):", "Audit Period", Format(Date, "mm/dd/yyyy"))
    
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "Invalid date format", vbExclamation
        Exit Sub
    End If
    
    ' Create audit file
    Set auditFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Documents\EmailAudit_" & Format(Now, "yyyymmdd") & ".csv")
    auditFile.WriteLine "Date,Sender,Recipients,Subject,Size,Importance,Sensitivity"
    
    ' Process sent items
    For Each olItem In olFolder.Items.Restrict("[SentOn] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' AND [SentOn] <= '" & Format(endDate + 1, "ddddd h:nn AMPM") & "'")
        If TypeName(olItem) = "MailItem" Then
            auditFile.WriteLine """" & olItem.SentOn & """,""" & olItem.SenderName & """,""" & olItem.To & """,""" & _
                              Replace(olItem.Subject, """", "'") & """," & olItem.Size & "," & _
                              olItem.Importance & "," & olItem.Sensitivity
        End If
    Next
    
    auditFile.Close
    MsgBox "Audit trail generated for sent emails between " & startDate & " and " & endDate, vbInformation
End Sub
```

## 8. Secure Email Expiration
```vba
Sub SetEmailExpiration()
    Dim olApp As Outlook.Application
    Dim olItem As Object
    
    Set olApp = Outlook.Application
    
    ' Check if an email is selected
    If TypeName(olApp.ActiveWindow) = "Inspector" Then
        Set olItem = olApp.ActiveInspector.CurrentItem
        
        If TypeName(olItem) = "MailItem" Then
            Dim expireDays As String
            expireDays = InputBox("Set expiration in days (message will be deleted after this period):", "Email Expiration", "7")
            
            If IsNumeric(expireDays) Then
                olItem.FlagRequest = "Expires in " & expireDays & " days"
                olItem.FlagDueBy = Date + CInt(expireDays)
                olItem.Save
                MsgBox "Expiration set for " & expireDays & " days from now.", vbInformation
            Else
                MsgBox "Please enter a valid number.", vbExclamation
            End If
        Else
            MsgBox "This is not an email message.", vbExclamation
        End If
    Else
        MsgBox "Please open an email first.", vbExclamation
    End If
End Sub
```

## 9. Compliance Checklist Verification
```vba
Sub ComplianceChecklist()
    Dim olApp As Outlook.Application
    Dim olTask As Outlook.TaskItem
    Dim checklist As String
    
    Set olApp = Outlook.Application
    Set olTask = olApp.CreateItem(olTaskItem)
    
    checklist = "COMPLIANCE CHECKLIST:" & vbCrLf & vbCrLf & _
               "1. Verify all security patches are installed" & vbCrLf & _
               "2. Review firewall rules" & vbCrLf & _
               "3. Check for unauthorized user accounts" & vbCrLf & _
               "4. Review privileged access" & vbCrLf & _
               "5. Verify backup integrity" & vbCrLf & _
               "6. Review incident reports" & vbCrLf & _
               "7. Check physical security controls" & vbCrLf & _
               "8. Review password policies" & vbCrLf & _
               "9. Audit log retention" & vbCrLf & _
               "10. Verify encryption standards"
    
    With olTask
        .Subject = "MONTHLY COMPLIANCE CHECKLIST - " & Format(Now, "mmmm yyyy")
        .Body = checklist & vbCrLf & vbCrLf & _
               "Completed By: _________________" & vbCrLf & _
               "Date: _________________" & vbCrLf & _
               "Approved By: _________________"
        .StartDate = Date
        .DueDate = DateSerial(Year(Date), Month(Date) + 1, 0) ' Last day of current month
        .ReminderSet = True
        .ReminderTime = DateSerial(Year(Date), Month(Date) + 1, 0) - 3 ' 3 days before due
        .Categories = "Compliance"
        .Status = olTaskNotStarted
        .Save
    End With
    
    MsgBox "Monthly compliance checklist task created.", vbInformation
End Sub
```

## 10. Secure Email Recipient Verification
```vba
Sub VerifyExternalRecipients()
    Dim olApp As Outlook.Application
    Dim olItem As Object
    Dim recip As Outlook.Recipient
    Dim externalDomains As Variant
    Dim externalFound As Boolean
    Dim confirmSend As VbMsgBoxResult
    
    Set olApp = Outlook.Application
    
    ' Check if an email is selected
    If TypeName(olApp.ActiveWindow) = "Inspector" Then
        Set olItem = olApp.ActiveInspector.CurrentItem
        
        If TypeName(olItem) = "MailItem" Then
            ' Define internal domains
            externalDomains = Array("@company.com", "@corp.company.com")
            externalFound = False
            
            ' Check all recipients
            For Each recip In olItem.Recipients
                Dim isInternal As Boolean
                isInternal = False
                
                ' Check if recipient matches any internal domain
                For Each domain In externalDomains
                    If InStr(1, recip.Address, domain, vbTextCompare) > 0 Then
                        isInternal = True
                        Exit For
                    End If
                Next
                
                If Not isInternal Then
                    externalFound = True
                    MsgBox "External recipient detected: " & recip.Name & " (" & recip.Address & ")", vbExclamation, "Security Alert"
                End If
            Next
            
            If externalFound Then
                confirmSend = MsgBox("This email contains external recipients." & vbCrLf & _
                                   "Are you sure you want to send it?", vbQuestion + vbYesNo, "Confirm Send")
                
                If confirmSend = vbNo Then
                    olItem.Close olDiscard
                End If
            End If
        End If
    End If
End Sub
```

## 11. Security Policy Acknowledgment Tracker
```vba
Sub TrackPolicyAcknowledgment()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim policyVersion As String
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    policyVersion = InputBox("Enter policy version number:", "Policy Acknowledgment", "v3.1")
    
    With olMail
        .Subject = "SECURITY POLICY ACKNOWLEDGMENT REQUIRED - " & policyVersion
        .HTMLBody = "<p>Dear Team,</p>" & _
                   "<p>Please acknowledge receipt and understanding of our updated Security Policy (" & policyVersion & ").</p>" & _
                   "<p><strong>Action Required:</strong> Reply to this email with 'I acknowledge' in the body.</p>" & _
                   "<p>The policy document is attached for your reference.</p>" & _
                   "<p>Thank you,<br>Security Team</p>"
        .To = "all-employees@company.com"
        .Attachments.Add "\\server\policies\SecurityPolicy_" & policyVersion & ".pdf"
        .Categories = "Compliance"
        .Importance = olImportanceHigh
        .FlagRequest = "Follow up"
        .FlagDueBy = Date + 14
        .Save
        
        ' Move to tracking folder
        Dim olNS As Outlook.NameSpace
        Dim olFolder As Outlook.MAPIFolder
        Set olNS = olApp.GetNamespace("MAPI")
        On Error Resume Next
        Set olFolder = olNS.GetDefaultFolder(olFolderInbox).Folders("PolicyAcks")
        If olFolder Is Nothing Then
            Set olFolder = olNS.GetDefaultFolder(olFolderInbox).Folders.Add("PolicyAcks")
        End If
        On Error GoTo 0
        
        .Move olFolder
    End With
    
    MsgBox "Policy acknowledgment request sent and saved for tracking.", vbInformation
End Sub
```

## 12. Automatic Email Encryption
```vba
Sub AutoEncryptSensitiveEmails()
    Dim olApp As Outlook.Application
    Dim olItem As Object
    Dim keywords As Variant
    Dim foundKeyword As Boolean
    
    Set olApp = Outlook.Application
    
    ' Check if an email is selected
    If TypeName(olApp.ActiveWindow) = "Inspector" Then
        Set olItem = olApp.ActiveInspector.CurrentItem
        
        If TypeName(olItem) = "MailItem" Then
            ' Define sensitive keywords
            keywords = Array("confidential", "secret", "proprietary", "ssn", "social security", _
                            "credit card", "password", "credentials", "restricted")
            
            ' Check subject and body for keywords
            foundKeyword = False
            For Each kw In keywords
                If InStr(1, olItem.Subject, kw, vbTextCompare) > 0 Or _
                   InStr(1, olItem.Body, kw, vbTextCompare) > 0 Then
                    foundKeyword = True
                    Exit For
                End If
            Next
            
            If foundKeyword Then
                If MsgBox("This email contains sensitive keywords. Encrypt it?", vbQuestion + vbYesNo, "Encrypt Email") = vbYes Then
                    olItem.Sensitivity = olConfidential
                    olItem.Subject = "[ENCRYPT] " & olItem.Subject
                    olItem.Save
                    MsgBox "Email marked for encryption. Please ensure your encryption system is properly configured.", vbInformation
                End If
            End If
        End If
    End If
End Sub
```

## 13. Security Training Reminder
```vba
Sub ScheduleSecurityTrainingReminders()
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    Dim i As Integer
    
    Set olApp = Outlook.Application
    
    ' Create quarterly training reminders
    For i = 0 To 3
        Set olAppt = olApp.CreateItem(olAppointmentItem)
        With olAppt
            .Subject = "MANDATORY SECURITY TRAINING - Q" & (i + 1)
            .Body = "This is your quarterly mandatory security training reminder." & vbCrLf & _
                   "Please complete the training at: https://training.company.com/security" & vbCrLf & vbCrLf & _
                   "Due by: " & Format(DateSerial(Year(Date), (i * 3) + 3, 15), "mmmm d, yyyy") & vbCrLf & _
                   "Time required: 30 minutes" & vbCrLf & _
                   "Contact security-training@company.com with questions"
            .Start = DateSerial(Year(Date), (i * 3) + 1, 1) ' First day of quarter
            .Duration = 30
            .ReminderSet = True
            .ReminderMinutesBeforeStart = 1440 ' 1 day before
            .Categories = "Training;Security"
            .BusyStatus = olBusy
            .Save
        End With
    Next
    
    MsgBox "Quarterly security training reminders scheduled for the year.", vbInformation
End Sub
```

## 14. Phishing Test Simulation
```vba
Sub RunPhishingTest()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim testType As Integer
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    testType = InputBox("Select phishing test type:" & vbCrLf & _
                       "1. Fake Password Reset" & vbCrLf & _
                       "2. Urgent Action Required" & vbCrLf & _
                       "3. Fake Invoice" & vbCrLf & _
                       "4. Fake Shipping Notification", "Phishing Test", "1")
    
    With olMail
        Select Case testType
            Case 1
                .Subject = "Urgent: Password Reset Required"
                .HTMLBody = "<p>Dear User,</p>" & _
                           "<p>Our system detected unusual activity on your account. You must reset your password immediately.</p>" & _
                           "<p><a href='http://fake.company.com/reset'>Click here to reset your password</a></p>" & _
                           "<p>If you didn't request this, please contact IT immediately.</p>" & _
                           "<p>IT Support Team</p>"
            Case 2
                .Subject = "Action Required: Your Account Will Be Suspended"
                .HTMLBody = "<p>Hello,</p>" & _
                           "<p>Your account will be suspended in 24 hours unless you verify your details.</p>" & _
                           "<p><a href='http://fake.company.com/verify'>Click here to verify your account</a></p>" & _
                           "<p>This is an automated message. Please do not reply.</p>"
            Case 3
                .Subject = "Invoice #INV-2023-456 Pending Payment"
                .HTMLBody = "<p>Attached is your invoice for recent services.</p>" & _
                           "<p>Amount Due: $249.99</p>" & _
                           "<p>Due Date: " & Format(Date + 7, "mmmm d, yyyy") & "</p>" & _
                           "<p><a href='http://fake.company.com/pay'>Click here to pay online</a></p>" & _
                           "<p>Questions? Call 1-800-FAKE-NUM</p>"
            Case 4
                .Subject = "Your Package Delivery Notification #DL-78945"
                .HTMLBody = "<p>Your package is scheduled for delivery tomorrow.</p>" & _
                           "<p>Tracking Number: DL-78945</p>" & _
                           "<p><a href='http://fake.company.com/track'>Track your package</a></p>" & _
                           "<p>If you didn't expect this delivery, click to report.</p>"
        End Select
        
        .To = "test-group@company.com"
        .Categories = "Security Test"
        .Save
        
        ' Move to phishing test folder
        Dim olNS As Outlook.NameSpace
        Dim olFolder As Outlook.MAPIFolder
        Set olNS = olApp.GetNamespace("MAPI")
        On Error Resume Next
        Set olFolder = olNS.GetDefaultFolder(olFolderInbox).Folders("PhishingTests")
        If olFolder Is Nothing Then
            Set olFolder = olNS.GetDefaultFolder(olFolderInbox).Folders.Add("PhishingTests")
        End If
        On Error GoTo 0
        
        .Move olFolder
    End With
    
    MsgBox "Phishing test email created and saved. Monitor for click-throughs.", vbInformation
End Sub
```

## 15. Secure File Transfer Notification
```vba
Sub SendSecureFileTransfer()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim filePath As String
    Dim fso As Object
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get file to transfer
    filePath = Application.GetOpenFilename("All Files (*.*), *.*")
    If filePath = "False" Then Exit Sub
    
    With olMail
        .Subject = "Secure File Transfer Notification - " & fso.GetFileName(filePath)
        .HTMLBody = "<p>Dear Recipient,</p>" & _
                   "<p>A secure file has been shared with you via our corporate secure file transfer system.</p>" & _
                   "<p><strong>File Name:</strong> " & fso.GetFileName(filePath) & "</p>" & _
                   "<p><strong>File Size:</strong> " & FormatNumber(fso.GetFile(filePath).Size / 1024, 1) & " KB</p>" & _
                   "<p><strong>Instructions:</strong></p>" & _
                   "<ol>" & _
                   "<li>Go to https://securetransfer.company.com</li>" & _
                   "<li>Log in with your corporate credentials</li>" & _
                   "<li>Navigate to 'My Files'</li>" & _
                   "<li>Download the file within the next 7 days</li>" & _
                   "</ol>" & _
                   "<p>This file was not attached directly to this email for security reasons.</p>" & _
                   "<p>If you have any issues accessing the file, please contact IT Support.</p>"
        .To = InputBox("Enter recipient email:", "Secure Transfer")
        If .To = "" Then Exit Sub
        .Categories = "Secure Transfer"
        .Importance = olImportanceHigh
        .Display
    End With
End Sub
```

## 16. Access Review Task Generator
```vba
Sub GenerateAccessReviewTasks()
    Dim olApp As Outlook.Application
    Dim olTask As Outlook.TaskItem
    Dim reviewers As Variant
    Dim i As Integer
    
    Set olApp = Outlook.Application
    reviewers = Array("admin1@company.com", "admin2@company.com", "security-team@company.com")
    
    For i = LBound(reviewers) To UBound(reviewers)
        Set olTask = olApp.CreateItem(olTaskItem)
        With olTask
            .Subject = "QUARTERLY ACCESS REVIEW - " & Format(Date, "Q\\QQ yyyy")
            .Body = "As part of our security compliance program, please review the following:" & vbCrLf & vbCrLf & _
                   "1. Review all admin accounts in your department" & vbCrLf & _
                   "2. Verify all active accounts belong to current employees" & vbCrLf & _
                   "3. Check for excessive permissions" & vbCrLf & _
                   "4. Document any changes made" & vbCrLf & vbCrLf & _
                   "Complete by: " & Format(Date + 14, "mmmm d, yyyy") & vbCrLf & _
                   "Submit report to: security-compliance@company.com"
            .Assign
            .Recipients.Add reviewers(i)
            .DueDate = Date + 14
            .ReminderSet = True
            .ReminderTime = Date + 7
            .Categories = "Access Review"
            .Status = olTaskWaiting
            .Save
        End With
    Next
    
    MsgBox "Access review tasks assigned to " & UBound(reviewers) + 1 & " reviewers.", vbInformation
End Sub
```

## 17. Secure Meeting Request
```vba
Sub CreateSecureMeeting()
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    
    Set olApp = Outlook.Application
    Set olAppt = olApp.CreateItem(olAppointmentItem)
    
    With olAppt
        .Subject = "SECURE: " & InputBox("Enter meeting subject:", "Secure Meeting")
        .Body = "This is a secure meeting to discuss confidential matters." & vbCrLf & vbCrLf & _
               "Location: " & InputBox("Enter secure location:", "Meeting Location", "Conference Room A") & vbCrLf & _
               "Agenda:" & vbCrLf & _
               "1. " & InputBox("Enter first agenda item:", "Agenda") & vbCrLf & _
               "2. " & InputBox("Enter second agenda item:", "Agenda") & vbCrLf & _
               "3. " & InputBox("Enter third agenda item:", "Agenda") & vbCrLf & vbCrLf & _
               "NOTE: No electronic devices permitted. Bring government-issued ID."
        .Start = CDate(InputBox("Enter meeting date/time (mm/dd/yyyy hh:mm AM/PM):", "Meeting Time", Format(Date + 1 & " 10:00 AM", "mm/dd/yyyy hh:mm AM/PM")))
        .Duration = 60
        .Sensitivity = olConfidential
        .Categories = "Secure Meeting"
        .MeetingStatus = olMeeting
        .RequiredAttendees = InputBox("Enter required attendees (comma separated):", "Attendees")
        .Display
    End With
End Sub
```

## 18. Security Bulletin Distribution
```vba
Sub DistributeSecurityBulletin()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim bulletinType As String
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    bulletinType = InputBox("Enter bulletin type:" & vbCrLf & _
                          "1. Critical" & vbCrLf & _
                          "2. High" & vbCrLf & _
                          "3. Medium" & vbCrLf & _
                          "4. Low", "Security Bulletin", "2")
    
    With olMail
        Select Case bulletinType
            Case "1"
                .Subject = "[CRITICAL] Security Bulletin: " & InputBox("Enter bulletin title:", "Bulletin Title")
                .Importance = olImportanceHigh
                .Sensitivity = olConfidential
            Case "2"
                .Subject = "[HIGH] Security Bulletin: " & InputBox("Enter bulletin title:", "Bulletin Title")
                .Importance = olImportanceHigh
            Case "3"
                .Subject = "[MEDIUM] Security Bulletin: " & InputBox("Enter bulletin title:", "Bulletin Title")
                .Importance = olImportanceNormal
            Case "4"
                .Subject = "[LOW] Security Bulletin: " & InputBox("Enter bulletin title:", "Bulletin Title")
                .Importance = olImportanceLow
        End Select
        
        .HTMLBody = "<p><strong>Security Bulletin: " & .Subject & "</strong></p>" & _
                   "<p><strong>Date:</strong> " & Format(Date, "mmmm d, yyyy") & "</p>" & _
                   "<p><strong>Affected Systems:</strong> " & InputBox("Enter affected systems:", "Affected Systems") & "</p>" & _
                   "<p><strong>Description:</strong></p>" & _
                   "<p>" & InputBox("Enter bulletin description:", "Description") & "</p>" & _
                   "<p><strong>Action Required:</strong></p>" & _
                   "<p>" & InputBox("Enter required actions:", "Actions") & "</p>" & _
                   "<p><strong>Contact:</strong> security-team@company.com</p>"
        
        ' Add appropriate distribution list based on severity
        Select Case bulletinType
            Case "1", "2"
                .To = "all-employees@company.com"
            Case "3"
                .To = "it-staff@company.com;department-heads@company.com"
            Case "4"
                .To = "it-security@company.com"
        End Select
        
        .Categories = "Security Bulletin"
        .Save
        
        ' Move to bulletins folder
        Dim olNS As Outlook.NameSpace
        Dim olFolder As Outlook.MAPIFolder
        Set olNS = olApp.GetNamespace("MAPI")
        On Error Resume Next
        Set olFolder = olNS.GetDefaultFolder(olFolderSentMail).Folders("SecurityBulletins")
        If olFolder Is Nothing Then
            Set olFolder = olNS.GetDefaultFolder(olFolderSentMail).Folders.Add("SecurityBulletins")
        End If
        On Error GoTo 0
        
        .Move olFolder
    End With
    
    MsgBox "Security bulletin prepared and saved. Ready for distribution.", vbInformation
End Sub
```

## 19. Data Retention Cleanup
```vba
Sub CleanupOldDataByPolicy()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim policy As String
    Dim cutoffDate As Date
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Select folder to clean up
    Set olFolder = olApp.Session.PickFolder
    If olFolder Is Nothing Then Exit Sub
    
    ' Get retention policy
    policy = InputBox("Select retention policy:" & vbCrLf & _
                     "1. Short-term (1 year)" & vbCrLf & _
                     "2. Medium-term (3 years)" & vbCrLf & _
                     "3. Long-term (7 years)", "Retention Policy", "1")
    
    Select Case policy
        Case "1"
            cutoffDate = DateAdd("yyyy", -1, Date)
        Case "2"
            cutoffDate = DateAdd("yyyy", -3, Date)
        Case "3"
            cutoffDate = DateAdd("yyyy", -7, Date)
        Case Else
            MsgBox "Invalid policy selected", vbExclamation
            Exit Sub
    End Select
    
    ' Process items older than cutoff
    count = 0
    For Each olItem In olFolder.Items.Restrict("[ReceivedTime] < '" & Format(cutoffDate, "ddddd h:nn AMPM") & "'")
        If TypeName(olItem) = "MailItem" Then
            ' Check for exceptions (flagged items, high importance, etc.)
            If olItem.IsMarkedAsTask = False And _
               olItem.Importance <> olImportanceHigh And _
               olItem.Sensitivity <> olConfidential Then
                olItem.Delete
                count = count + 1
            End If
        End If
    Next
    
    MsgBox "Deleted " & count & " items older than " & cutoffDate & " per retention policy.", vbInformation
End Sub
```

## 20. Secure Contact Verification
```vba
Sub VerifyExternalContacts()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim contactFile As Object
    Dim fso As Object
    Dim externalDomains As Variant
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderContacts)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define internal domains
    externalDomains = Array("@company.com", "@corp.company.com")
    
    ' Create verification report
    Set contactFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Documents\ExternalContacts_" & Format(Now, "yyyymmdd") & ".csv")
    contactFile.WriteLine "Name,Email,Phone,LastContact,Notes"
    
    ' Process all contacts
    count = 0
    For Each olItem In olFolder.Items
        If TypeName(olItem) = "ContactItem" Then
            Dim isExternal As Boolean
            isExternal = True
            
            ' Check email addresses
            If olItem.Email1Address <> "" Then
                For Each domain In externalDomains
                    If InStr(1, olItem.Email1Address, domain, vbTextCompare) > 0 Then
                        isExternal = False
                        Exit For
                    End If
                Next
            End If
            
            ' Check secondary email
            If isExternal And olItem.Email2Address <> "" Then
                For Each domain In externalDomains
                    If InStr(1, olItem.Email2Address, domain, vbTextCompare) > 0 Then
                        isExternal = False
                        Exit For
                    End If
                Next
            End If
            
            ' Check tertiary email
            If isExternal And olItem.Email3Address <> "" Then
                For Each domain In externalDomains
                    If InStr(1, olItem.Email3Address, domain, vbTextCompare) > 0 Then
                        isExternal = False
                        Exit For
                    End If
                Next
            End If
            
            ' Add to report if external
            If isExternal Then
                contactFile.WriteLine """" & olItem.FullName & """,""" & olItem.Email1Address & """,""" & _
                                    olItem.BusinessTelephoneNumber & """,""" & olItem.LastModificationTime & """,""" & _
                                    "Verify this contact is still needed"""
                count = count + 1
            End If
        End If
    Next
    
    contactFile.Close
    MsgBox "Identified " & count & " external contacts. Report saved to:" & vbCrLf & _
           Environ("USERPROFILE") & "\Documents\ExternalContacts_" & Format(Now, "yyyymmdd") & ".csv", vbInformation
End Sub
```

## 21. Secure Note Creator
```vba
Sub CreateSecureNote()
    Dim olApp As Outlook.Application
    Dim olNote As Outlook.NoteItem
    
    Set olApp = Outlook.Application
    Set olNote = olApp.CreateItem(olNoteItem)
    
    With olNote
        .Body = InputBox("Enter your secure note:", "Secure Note") & vbCrLf & vbCrLf & _
               "Created: " & Now & vbCrLf & _
               "By: " & olApp.Session.CurrentUser.Name
        .Categories = "Secure Note"
        .Color = olNoteColorYellow
        .Save
        
        ' Move to secure notes folder
        Dim olNS As Outlook.NameSpace
        Dim olFolder As Outlook.MAPIFolder
        Set olNS = olApp.GetNamespace("MAPI")
        On Error Resume Next
        Set olFolder = olNS.GetDefaultFolder(olFolderNotes).Folders("Secure Notes")
        If olFolder Is Nothing Then
            Set olFolder = olNS.GetDefaultFolder(olFolderNotes).Folders.Add("Secure Notes")
        End If
        On Error GoTo 0
        
        .Move olFolder
    End With
    
    MsgBox "Secure note created and stored.", vbInformation
End Sub
```

## 22. Security Exception Request
```vba
Sub RequestSecurityException()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .Subject = "SECURITY EXCEPTION REQUEST - " & InputBox("Briefly describe the exception:", "Exception Description")
        .HTMLBody = "<p><strong>Security Exception Request</strong></p>" & _
                   "<p><strong>Requester:</strong> " & olApp.Session.CurrentUser.Name & "</p>" & _
                   "<p><strong>Date:</strong> " & Format(Date, "mmmm d, yyyy") & "</p>" & _
                   "<p><strong>Exception Details:</strong></p>" & _
                   "<p>" & InputBox("Describe the exception in detail:", "Details") & "</p>" & _
                   "<p><strong>Business Justification:</strong></p>" & _
                   "<p>" & InputBox("Provide business justification:", "Justification") & "</p>" & _
                   "<p><strong>Proposed Mitigation:</strong></p>" & _
                   "<p>" & InputBox("Describe any risk mitigation:", "Mitigation") & "</p>" & _
                   "<p><strong>Requested Duration:</strong> " & InputBox("Duration of exception (e.g., 30 days, permanent):", "Duration") & "</p>"
        .To = "security-exceptions@company.com"
        .CC = InputBox("CC (manager or other stakeholders):", "CC")
        .Categories = "Security Exception"
        .Importance = olImportanceHigh
        .Display
    End With
End Sub
```

## 23. Secure Calendar Review
```vba
Sub ReviewCalendarForSecurityIssues()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim reportFile As Object
    Dim fso As Object
    Dim startDate As Date
    Dim endDate As Date
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get date range
    startDate = InputBox("Enter start date (mm/dd/yyyy):", "Review Period", Format(Date - 30, "mm/dd/yyyy"))
    endDate = InputBox("Enter end date (mm/dd/yyyy):", "Review Period", Format(Date, "mm/dd/yyyy"))
    
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "Invalid date format", vbExclamation
        Exit Sub
    End If
    
    ' Create report file
    Set reportFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Documents\CalendarSecurityReview_" & Format(Now, "yyyymmdd") & ".csv")
    reportFile.WriteLine "Date,Subject,Location,Sensitivity,Confidential Items Found"
    
    ' Process calendar items
    count = 0
    For Each olItem In olFolder.Items.Restrict("[Start] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' AND [End] <= '" & Format(endDate + 1, "ddddd h:nn AMPM") & "'")
        If TypeName(olItem) = "AppointmentItem" Then
            Dim confidentialFound As Boolean
            confidentialFound = False
            
            ' Check for sensitive information
            If olItem.Sensitivity = olConfidential Then
                confidentialFound = True
                count = count + 1
            ElseIf InStr(1, olItem.Subject, "password", vbTextCompare) > 0 Or _
                   InStr(1, olItem.Subject, "secret", vbTextCompare) > 0 Or _
                   InStr(1, olItem.Body, "password", vbTextCompare) > 0 Or _
                   InStr(1, olItem.Body, "secret", vbTextCompare) > 0 Then
                confidentialFound = True
                count = count + 1
            End If
            
            reportFile.WriteLine """" & olItem.Start & """,""" & olItem.Subject & """,""" & _
                               olItem.Location & """,""" & olItem.Sensitivity & """,""" & _
                               IIf(confidentialFound, "YES", "NO") & """"
        End If
    Next
    
    reportFile.Close
    MsgBox "Calendar review complete. Found " & count & " potentially confidential items." & vbCrLf & _
           "Report saved to:" & vbCrLf & _
           Environ("USERPROFILE") & "\Documents\CalendarSecurityReview_" & Format(Now, "yyyymmdd") & ".csv", vbInformation
End Sub
```

## 24. Secure Email Template
```vba
Sub CreateSecureEmailTemplate()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .Subject = "[SECURE] " & InputBox("Enter email subject:", "Secure Email")
        .HTMLBody = "<p><strong>SECURE COMMUNICATION</strong></p>" & _
                   "<p>This email contains confidential information intended only for the recipient(s) named below.</p>" & _
                   "<p>If you are not the intended recipient, please:" & _
                   "<ol>" & _
                   "<li>Do not read, copy, or forward this message</li>" & _
                   "<li>Notify the sender immediately by replying to this email</li>" & _
                   "<li>Delete this message from your system</li>" & _
                   "</ol>" & _
                   "<p>---------------------------------------</p>" & _
                   "<p>" & InputBox("Enter your message:", "Email Content") & "</p>" & _
                   "<p>---------------------------------------</p>" & _
                   "<p><strong>CONFIDENTIALITY NOTICE:</strong> This message and any attachments are intended only for the use of the individual or entity to which they are addressed and may contain information that is privileged, confidential, and exempt from disclosure under applicable law. If you are not the intended recipient, you are hereby notified that any dissemination, distribution, or copying of this communication is strictly prohibited.</p>"
        .To = InputBox("Enter recipient(s):", "Recipients")
        .CC = InputBox("Enter CC recipient(s) if any:", "CC")
        .Sensitivity = olConfidential
        .Categories = "Secure Communication"
        .Importance = olImportanceHigh
        .Display
    End With
End Sub
```

## 25. Security Log Export
```vba
Sub ExportSecurityLogs()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim logFile As Object
    Dim fso As Object
    Dim logType As String
    Dim count As Integer
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Select log type
    logType = InputBox("Select log type to export:" & vbCrLf & _
                      "1. Security Alerts" & vbCrLf & _
                      "2. Policy Violations" & vbCrLf & _
                      "3. Access Requests", "Log Export", "1")
    
    ' Select folder containing logs
    Set olFolder = olApp.Session.PickFolder
    If olFolder Is Nothing Then Exit Sub
    
    ' Create log file
    Set logFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Documents\SecurityLogs_" & Format(Now, "yyyymmdd") & ".csv")
    logFile.WriteLine "Date,Subject,Category,Status,Notes"
    
    ' Process items
    count = 0
    For Each olItem In olFolder.Items
        If TypeName(olItem) = "MailItem" Then
            ' Filter by type if needed
            If (logType = "1" And InStr(1, olItem.Subject, "Alert", vbTextCompare) > 0) Or _
               (logType = "2" And InStr(1, olItem.Subject, "Violation", vbTextCompare) > 0) Or _
               (logType = "3" And InStr(1, olItem.Subject, "Request", vbTextCompare) > 0) Or _
               logType = "" Then
                
                logFile.WriteLine """" & olItem.ReceivedTime & """,""" & olItem.Subject & """,""" & _
                                olItem.Categories & """,""" & IIf(olItem.UnRead, "New", "Processed") & """,""" & _
                                Left(Replace(olItem.Body, vbCrLf, " "), 100) & """"
                count = count + 1
            End If
        End If
    Next
    
    logFile.Close
    MsgBox "Exported " & count & " log entries to:" & vbCrLf & _
           Environ("USERPROFILE") & "\Documents\SecurityLogs_" & Format(Now, "yyyymmdd") & ".csv", vbInformation
End Sub
```

## 26. Secure Document Request
```vba
Sub RequestSecureDocument()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    Set olApp = Outlook.Application
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .Subject = "SECURE DOCUMENT REQUEST - " & InputBox("Document title:", "Document Request")
        .HTMLBody = "<p><strong>Secure Document Request</strong></p>" & _
                   "<p><strong>Requester:</strong> " & olApp.Session.CurrentUser.Name & "</p>" & _
                   "<p><strong>Department:</strong> " & InputBox("Your department:", "Department") & "</p>" & _
                   "<p><strong>Document Title:</strong> " & .Subject & "</p>" & _
                   "<p><strong>Purpose:</strong> " & InputBox("Purpose of request:", "Purpose") & "</p>" & _
                   "<p><strong>Required By:</strong> " & InputBox("Required date (mm/dd/yyyy):", "Due Date", Format(Date + 7, "mm/dd/yyyy")) & "</p>" & _
                   "<p><strong>Classification Level Needed:</strong> " & _
                   InputBox("Classification level:" & vbCrLf & _
                          "1. Internal" & vbCrLf & _
                          "2. Confidential" & vbCrLf & _
                          "3. Restricted", "Classification", "2") & "</p>" & _
                   "<p><strong>Delivery Method:</strong> " & _
                   InputBox("Preferred delivery method:" & vbCrLf & _
                          "1. Secure Email" & vbCrLf & _
                          "2. Encrypted USB" & vbCrLf & _
                          "3. Secure File Transfer", "Delivery", "1") & "</p>"
        .To = "document-control@company.com"
        .CC = InputBox("CC (manager or other stakeholders):", "CC")
        .Categories = "Document Request"
        .Importance = olImportanceHigh
        .Display
    End With
End Sub
```

