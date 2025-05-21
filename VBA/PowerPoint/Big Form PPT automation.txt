VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9660.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm Code
Private Sub cmdAddTopic_Click()
    ' Add the topic name from txtTopicName to the ListBox
    If Trim(txtTopicName.Text) <> "" Then
        lstTopics.AddItem txtTopicName.Text
        txtTopicName.Text = "" ' Clear the textbox after adding
    Else
        MsgBox "Please enter a valid topic name!", vbExclamation
    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim numSlides As Integer
    Dim mainTitle As String
    Dim numTopics As Integer
    Dim topicTitles() As String
    Dim i As Integer

    ' Validate inputs
    If IsNumeric(txtNumSlides.Text) And CInt(txtNumSlides.Text) >= 2 Then
        numSlides = CInt(txtNumSlides.Text)
    Else
        MsgBox "You must enter a valid number of slides (minimum 2)!", vbExclamation
        Exit Sub
    End If

    If Trim(txtMainTitle.Text) = "" Then
        MsgBox "Please enter a presentation title!", vbExclamation
        Exit Sub
    Else
        mainTitle = txtMainTitle.Text
    End If

    If lstTopics.ListCount = 0 Then
        MsgBox "Please add at least one topic to the Index Slide!", vbExclamation
        Exit Sub
    Else
        numTopics = lstTopics.ListCount
        ReDim topicTitles(1 To numTopics)
        For i = 1 To numTopics
            topicTitles(i) = lstTopics.List(i - 1)
        Next i
    End If

    ' Call the main PowerPoint generation subroutine
    Call GeneratePowerPoint(numSlides, mainTitle, numTopics, topicTitles)

    ' Close the UserForm
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ' Close the UserForm without generating the presentation
    Unload Me
End Sub

Sub ShowUserForm()
    ' Display the UserForm
    UserForm1.Show
End Sub

Sub GeneratePowerPoint(numSlides As Integer, mainTitle As String, numTopics As Integer, topicTitles() As String)
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slide As Object
    Dim slideIndex As Integer
    Dim savePath As String
    Dim customTitle As String
    Dim customText As String
    Dim i As Integer

    ' Start PowerPoint application
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application") ' Check if PowerPoint is already open
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application") ' Open PowerPoint if not already open
    End If
    On Error GoTo 0

    ' Make PowerPoint visible
    pptApp.Visible = True

    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add

    ' Add the Title Slide
    Const ppLayoutTitle As Long = 1 ' Layout index for "Title Slide"
    Set slide = pptPres.Slides.Add(1, ppLayoutTitle)
    slide.Shapes(1).TextFrame.TextRange.Text = mainTitle
    slide.Shapes(2).TextFrame.TextRange.Text = "Created with VBA Automation"

    ' Add the Index Slide
    Const ppLayoutText As Long = 2 ' Layout index for "Title and Text"
    Set slide = pptPres.Slides.Add(2, ppLayoutText)
    slide.Shapes(1).TextFrame.TextRange.Text = "Index"

    ' Add the topics as a bulleted list
    Dim topicList As String
    topicList = ""
    For i = 1 To numTopics
        topicList = topicList & i & ". " & topicTitles(i) & vbCrLf
    Next i
    slide.Shapes(2).TextFrame.TextRange.Text = topicList

    ' Add the remaining slides (Title and Text)
    For slideIndex = 3 To numSlides
        ' Add a new slide with the "Title and Text" layout
        Set slide = pptPres.Slides.Add(slideIndex, ppLayoutText)

        ' Ask the user for the title and text for each slide
        customTitle = InputBox("Enter the title for Slide " & slideIndex & ":", "Slide Title")
        customText = InputBox("Enter the text content for Slide " & slideIndex & ":", "Slide Content")

        ' Add the custom title and text to the slide
        slide.Shapes(1).TextFrame.TextRange.Text = customTitle
        slide.Shapes(2).TextFrame.TextRange.Text = customText
    Next slideIndex

    ' Prompt the user to save the presentation
    savePath = InputBox("Enter the full path to save the presentation (e.g., C:\MyPresentation.pptx):", "Save Presentation")
    
    ' Validate the save path
    If savePath = "" Then
        MsgBox "Save canceled. Please save the presentation manually in PowerPoint.", vbInformation
    Else
        ' Save the presentation
        pptPres.SaveAs savePath
        MsgBox "Presentation saved successfully to: " & savePath, vbInformation
    End If

    ' Release objects
    Set slide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub


