' Standard Module Code
Sub ShowPresentationGenerator()
    ' Display the UserForm
    UserForm1.Show
End Sub

Sub GeneratePowerPoint(numSlides As Integer, mainTitle As String, numTopics As Integer, topicTitles() As String)
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slide As Object
    Dim slideIndex As Integer
    Dim savePath As String
    Dim fileFormat As Long
    Dim customTitle As String
    Dim customText As String
    Dim i As Integer
    Dim fd As FileDialog
    
    On Error GoTo ErrorHandler
    
    ' Start PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Add the Title Slide
    Set slide = pptPres.Slides.Add(1, 1) ' ppLayoutTitle = 1
    slide.Shapes(1).TextFrame.TextRange.Text = mainTitle
    slide.Shapes(2).TextFrame.TextRange.Text = "Created with VBA Automation"
    
    ' Add the Index Slide
    Set slide = pptPres.Slides.Add(2, 2) ' ppLayoutText = 2
    slide.Shapes(1).TextFrame.TextRange.Text = "Index"
    
    ' Add the topics as a bulleted list
    Dim topicList As String
    topicList = ""
    For i = 1 To numTopics
        topicList = topicList & i & ". " & topicTitles(i) & vbCrLf
    Next i
    slide.Shapes(2).TextFrame.TextRange.Text = topicList
    
    ' Add the remaining slides (Title and Text)
    For slideIndex = 3 To numSlides + 2 ' +2 for title and index slides
        Set slide = pptPres.Slides.Add(slideIndex, 2) ' ppLayoutText = 2
        
        ' Get slide content from user
        customTitle = InputBox("Enter the title for Slide " & slideIndex & ":", "Slide Title", "Slide " & slideIndex)
        If customTitle = "" Then customTitle = "Slide " & slideIndex
        
        customText = InputBox("Enter the text content for Slide " & slideIndex & ":", "Slide Content", "Content for slide " & slideIndex)
        If customText = "" Then customText = "Add your content here"
        
        ' Add content to slide
        slide.Shapes(1).TextFrame.TextRange.Text = customTitle
        slide.Shapes(2).TextFrame.TextRange.Text = customText
    Next slideIndex
    
    ' Prompt user to save the presentation with FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "Save Presentation As"
        .InitialFileName = mainTitle & ".pptx"
        .Filters.Clear
        .Filters.Add "PowerPoint Presentation", "*.pptx"
        .Filters.Add "PowerPoint Macro-Enabled Presentation", "*.pptm"
        .Filters.Add "PowerPoint 97-2003 Presentation", "*.ppt"
        
        If .Show = -1 Then ' If user clicked Save
            savePath = .SelectedItems(1)
            
            ' Determine file format based on extension
            Select Case LCase(Right(savePath, 5))
                Case ".pptx": fileFormat = 24 ' ppSaveAsOpenXMLPresentation
                Case ".pptm": fileFormat = 25 ' ppSaveAsOpenXMLPresentationMacroEnabled
                Case ".ppt": fileFormat = 1  ' ppSaveAsPresentation
                Case Else: savePath = savePath & ".pptx": fileFormat = 24
            End Select
            
            ' Save the presentation
            pptPres.SaveAs savePath, fileFormat
            MsgBox "Presentation saved successfully to:" & vbCrLf & savePath, vbInformation
        Else
            MsgBox "Save canceled. Please save the presentation manually in PowerPoint.", vbExclamation
        End If
    End With
    
Cleanup:
    ' Release objects
    On Error Resume Next
    Set slide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set fd = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred:" & vbCrLf & Err.Description, vbCritical
    Resume Cleanup
End Sub
