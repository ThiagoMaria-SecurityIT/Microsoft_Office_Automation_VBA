' UserForm Code
Private Sub UserForm_Initialize()
    ' Set default number of slides
    txtNumSlides.Text = "4"
    
    ' Set form caption
    Me.Caption = "PowerPoint Presentation Generator"
End Sub

Private Sub cmdAddTopic_Click()
    ' Add the topic name from txtTopicName to the ListBox
    If Trim(txtTopicName.Text) <> "" Then
        lstTopics.AddItem txtTopicName.Text
        txtTopicName.Text = "" ' Clear the textbox after adding
        txtTopicName.SetFocus ' Return focus to input box
    Else
        MsgBox "Please enter a valid topic name!", vbExclamation
    End If
End Sub

Private Sub cmdRemoveTopic_Click()
    ' Remove selected topic from the ListBox
    If lstTopics.ListIndex <> -1 Then ' Check if an item is selected
        lstTopics.RemoveItem lstTopics.ListIndex
    Else
        MsgBox "Please select a topic to remove!", vbExclamation
    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim numSlides As Integer
    Dim mainTitle As String
    Dim numTopics As Integer
    Dim topicTitles() As String
    Dim i As Integer
    
    ' Validate main title
    If Trim(txtMainTitle.Text) = "" Then
        MsgBox "Please enter a presentation title!", vbExclamation
        txtMainTitle.SetFocus
        Exit Sub
    Else
        mainTitle = txtMainTitle.Text
    End If
    
    ' Validate number of slides
    If Not IsNumeric(txtNumSlides.Text) Then
        MsgBox "Number of slides must be a numeric value!", vbExclamation
        txtNumSlides.SetFocus
        Exit Sub
    End If
    
    numSlides = Val(txtNumSlides.Text)
    
    If numSlides < 2 Then
        MsgBox "Minimum of 2 slides required!", vbExclamation
        txtNumSlides.SetFocus
        Exit Sub
    End If
    
    ' Validate topics
    If lstTopics.ListCount = 0 Then
        MsgBox "Please add at least one topic!", vbExclamation
        txtTopicName.SetFocus
        Exit Sub
    Else
        numTopics = lstTopics.ListCount
        ReDim topicTitles(1 To numTopics)
        For i = 1 To numTopics
            topicTitles(i) = lstTopics.List(i - 1)
        Next i
    End If
    
    ' Disable form while processing
    Me.Hide
    DoEvents
    
    ' Call the main PowerPoint generation subroutine
    Call GeneratePowerPoint(numSlides, mainTitle, numTopics, topicTitles)
    
    ' Close the UserForm
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ' Close the UserForm without generating the presentation
    Unload Me
End Sub
