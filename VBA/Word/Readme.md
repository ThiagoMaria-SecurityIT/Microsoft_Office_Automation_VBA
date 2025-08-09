# VBA Macros for Microsoft Word - For Security Information Analyst

A collection of 30 VBA macros designed specifically for security information analysts using Microsoft Word. These macros enhance document security, automate security-related tasks, and streamline security reporting workflows.

## Table of Contents
- [Document Security & Protection](#document-security--protection)
  - [1. Password Protect Document](#1-password-protect-document)
  - [2. Remove All Metadata](#2-remove-all-metadata)
  - [3. Redact Selected Text](#3-redact-selected-text)
  - [4. Sanitize Document](#4-sanitize-document)
  - [5. Encrypt Document Content](#5-encrypt-document-content)
- [Content Analysis & Processing](#content-analysis--processing)
  - [6. Find and Highlight Security Keywords](#6-find-and-highlight-security-keywords)
  - [7. Find and Replace Sensitive Information Patterns](#7-find-and-replace-sensitive-information-patterns)
  - [8. Extract All Hyperlinks](#8-extract-all-hyperlinks)
  - [9. Auto-classify Document Based on Content](#9-auto-classify-document-based-on-content)
  - [10. Generate Document Hash](#10-generate-document-hash)
- [Document Formatting & Templates](#document-formatting--templates)
  - [11. Add Watermark Based on Classification](#11-add-watermark-based-on-classification)
  - [12. Create Security Header/Footer](#12-create-security-headerfooter)
  - [13. Create Security Incident Report Template](#13-create-security-incident-report-template)
  - [14. Create Security Assessment Template](#14-create-security-assessment-template)
  - [15. Auto-number Security Findings](#15-auto-number-security-findings)
  - [16. Generate Table of Contents for Security Reports](#16-generate-table-of-contents-for-security-reports)
- [Document Analysis & Reporting](#document-analysis--reporting)
  - [17. Generate Document Statistics Report](#17-generate-document-statistics-report)
  - [18. Track Changes Summary](#18-track-changes-summary)
  - [19. Extract Comments for Review](#19-extract-comments-for-review)
  - [20. Check Document for Embedded Objects](#20-check-document-for-embedded-objects)
  - [21. Check Document for Accessibility Issues](#21-check-document-for-accessibility-issues)
- [Document Conversion & Export](#document-conversion--export)
  - [22. Export Document to PDF with Security](#22-export-document-to-pdf-with-security)
  - [23. Convert Document to Plain Text](#23-convert-document-to-plain-text)
  - [24. Export Document Properties to CSV](#24-export-document-properties-to-csv)
- [Document Management & Automation](#document-management--automation)
  - [25. Backup Current Document](#25-backup-current-document)
  - [26. Auto-save Document with Timestamp](#26-auto-save-document-with-timestamp)
  - [27. Log Document Access](#27-log-document-access)
  - [28. Batch Find and Replace Across Multiple Documents](#28-batch-find-and-replace-across-multiple-documents)
  - [29. Compare Document Versions](#29-compare-document-versions)
  - [30. Add Digital Signature](#30-add-digital-signature)

## Document Security & Protection

### 1. Password Protect Document
Adds password protection to the current document.

```vba
Sub PasswordProtectDocument()
    Dim password As String
    password = InputBox("Enter password to protect the document:")
    If password <> "" Then
        ActiveDocument.Password = password
        ActiveDocument.Save
        MsgBox "Document is now password protected."
    Else
        MsgBox "No password entered. Document remains unprotected."
    End If
End Sub
```

### 2. Remove All Metadata
Removes all document metadata and properties.

```vba
Sub RemoveAllMetadata()
    With ActiveDocument
        .BuiltInDocumentProperties("Title").Value = ""
        .BuiltInDocumentProperties("Subject").Value = ""
        .BuiltInDocumentProperties("Author").Value = ""
        .BuiltInDocumentProperties("Keywords").Value = ""
        .BuiltInDocumentProperties("Comments").Value = ""
    End With
    
    ActiveDocument.RemoveDocumentInformation (wdRDIAll)
    MsgBox "All metadata has been removed."
End Sub
```

### 3. Redact Selected Text
Redacts selected text by making it black with black highlight.

```vba
Sub RedactSelectedText()
    If Selection.Type <> wdSelectionIP Then
        Selection.Font.Color = wdColorBlack
        Selection.HighlightColorIndex = wdBlack
    Else
        MsgBox "Please select text to redact."
    End If
End Sub
```

### 4. Sanitize Document
Removes comments, revisions, hidden text, fields, headers/footers, and metadata.

```vba
Sub SanitizeDocument()
    ' Remove comments
    Dim comment As Comment
    For Each comment In ActiveDocument.Comments
        comment.Delete
    Next comment
    
    ' Remove revisions
    If ActiveDocument.Revisions.Count > 0 Then
        ActiveDocument.Revisions.AcceptAll
    End If
    
    ' Remove hidden text
    With ActiveDocument.Content.Find
        .Text = ""
        .Font.Hidden = True
        .Replacement.Text = ""
        .Replacement.Font.Hidden = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Remove document properties
    ActiveDocument.RemoveDocumentInformation (wdRDIAll)
    
    ' Remove all fields
    Dim field As Field
    For Each field In ActiveDocument.Fields
        field.Unlink
    Next field
    
    ' Remove all headers and footers
    Dim sec As Section
    Dim hdrftr As HeaderFooter
    
    For Each sec In ActiveDocument.Sections
        For Each hdrftr In sec.Headers
            hdrftr.Range.Delete
        Next hdrftr
        
        For Each hdrftr In sec.Footers
            hdrftr.Range.Delete
        Next hdrftr
    Next sec
    
    MsgBox "Document has been sanitized."
End Sub
```

### 5. Encrypt Document Content
Encrypts the document with a password.

```vba
Sub EncryptDocumentContent()
    Dim password As String
    password = InputBox("Enter password to encrypt the document:")
    
    If password <> "" Then
        ActiveDocument.Password = password
        ActiveDocument.WritePassword = password
        ActiveDocument.Save
        MsgBox "Document content has been encrypted."
    Else
        MsgBox "No password entered. Document remains unencrypted."
    End If
End Sub
```

## Content Analysis & Processing

### 6. Find and Highlight Security Keywords
Highlights security-related keywords in the document.

```vba
Sub HighlightSecurityKeywords()
    Dim keywords() As Variant
    keywords = Array("password", "confidential", "secret", "classified", "sensitive", "security", "vulnerability", "threat", "risk", "breach")
    
    Dim keyword As Variant
    For Each keyword In keywords
        With Selection.Find
            .Text = keyword
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute
                Selection.Range.HighlightColorIndex = wdYellow
            Loop
        End With
    Next keyword
    
    MsgBox "Security keywords have been highlighted."
End Sub
```

### 7. Find and Replace Sensitive Information Patterns
Replaces sensitive information patterns with [REDACTED].

```vba
Sub FindAndReplaceSensitiveInfo()
    ' Patterns for common sensitive information
    Dim patterns() As Variant
    patterns = Array("\d{3}-\d{2}-\d{4}", "SSN", "social security", "credit card", "account number", "password", "confidential")
    
    Dim i As Integer
    For i = LBound(patterns) To UBound(patterns)
        With Selection.Find
            .Text = patterns(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = (i = 0) ' Use wildcards only for SSN pattern
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute
                Selection.Range.Text = "[REDACTED]"
            Loop
        End With
    Next i
    
    MsgBox "Sensitive information patterns have been replaced with [REDACTED]."
End Sub
```

### 8. Extract All Hyperlinks
Extracts all hyperlinks to a new document.

```vba
Sub ExtractAllHyperlinks()
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    Dim hyperlink As Hyperlink
    Dim count As Integer
    
    count = 0
    newDoc.Content.Text = "Extracted Hyperlinks:" & vbCrLf & vbCrLf
    
    For Each hyperlink In ActiveDocument.Hyperlinks
        count = count + 1
        newDoc.Content.InsertAfter count & ". " & hyperlink.TextToDisplay & " - " & hyperlink.Address & vbCrLf
    Next hyperlink
    
    If count = 0 Then
        newDoc.Content.InsertAfter "No hyperlinks found in the document."
    End If
    
    MsgBox count & " hyperlinks have been extracted."
End Sub
```

### 9. Auto-classify Document Based on Content
Automatically classifies the document based on content keywords.

```vba
Sub AutoClassifyDocument()
    Dim content As String
    content = LCase(ActiveDocument.Content.Text)
    
    Dim classification As String
    Dim confidence As Integer
    
    ' Keywords for different classification levels
    Dim publicKeywords() As Variant
    publicKeywords = Array("public", "for release", "unclassified", "open source")
    
    Dim confidentialKeywords() As Variant
    confidentialKeywords = Array("confidential", "internal use", "company private", "restricted")
    
    Dim secretKeywords() As Variant
    secretKeywords = Array("secret", "top secret", "classified", "sensitive")
    
    ' Count occurrences of keywords
    Dim publicCount As Integer, confidentialCount As Integer, secretCount As Integer
    Dim keyword As Variant
    
    For Each keyword In publicKeywords
        publicCount = publicCount + CountOccurrences(content, keyword)
    Next keyword
    
    For Each keyword In confidentialKeywords
        confidentialCount = confidentialCount + CountOccurrences(content, keyword)
    Next keyword
    
    For Each keyword In secretKeywords
        secretCount = secretCount + CountOccurrences(content, keyword)
    Next keyword
    
    ' Determine classification based on keyword counts
    If secretCount > 0 Then
        classification = "SECRET"
        confidence = secretCount
    ElseIf confidentialCount > 0 Then
        classification = "CONFIDENTIAL"
        confidence = confidentialCount
    ElseIf publicCount > 0 Then
        classification = "PUBLIC"
        confidence = publicCount
    Else
        classification = "UNCLASSIFIED"
        confidence = 0
    End If
    
    ' Display classification result
    Dim result As String
    result = "Document Classification: " & classification & vbCrLf
    result = result & "Confidence Level: " & confidence & " keyword matches"
    
    MsgBox result, vbInformation, "Classification Result"
    
    ' Add classification to document properties
    ActiveDocument.BuiltInDocumentProperties("Keywords").Value = classification
End Sub

' Helper function to count occurrences
Function CountOccurrences(text As String, word As String) As Integer
    Dim count As Integer
    Dim position As Integer
    
    count = 0
    position = InStr(1, text, word)
    
    Do While position > 0
        count = count + 1
        position = InStr(position + 1, text, word)
    Loop
    
    CountOccurrences = count
End Function
```

### 10. Generate Document Hash
Generates a simple hash of the document content.

```vba
Sub GenerateDocumentHash()
    ' This is a simplified hash function for demonstration
    Dim content As String
    content = ActiveDocument.Content.Text
    
    ' Simple hash function (not cryptographically secure)
    Dim hash As Long
    Dim i As Integer
    
    hash = 5381
    
    For i = 1 To Len(content)
        hash = hash * 33 + Asc(Mid(content, i, 1))
    Next i
    
    ' Convert to hexadecimal string
    Dim hexHash As String
    hexHash = Hex(hash)
    
    ' Display the hash
    Dim result As String
    result = "Document Hash: " & hexHash & vbCrLf & vbCrLf
    result = result & "Note: This is a simple hash function for demonstration purposes."
    
    MsgBox result, vbInformation, "Document Hash"
    
    ' Save the hash to document properties
    ActiveDocument.BuiltInDocumentProperties("Comments").Value = "Document Hash: " & hexHash
    ActiveDocument.Save
End Sub
```

## Document Formatting & Templates

### 11. Add Watermark Based on Classification
Adds a classification watermark to the document.

```vba
Sub AddClassificationWatermark()
    Dim classification As String
    classification = InputBox("Enter document classification (e.g., Public, Confidential, Secret):")
    
    If classification <> "" Then
        Dim watermarkColor As WdColor
        
        Select Case LCase(classification)
            Case "public"
                watermarkColor = wdColorBlue
            Case "confidential"
                watermarkColor = wdColorGreen
            Case "secret"
                watermarkColor = wdColorRed
            Case Else
                watermarkColor = wdColorBlack
        End Select
        
        ActiveDocument.Sections(1).Range.Select
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        
        Selection.HeaderFooter.Shapes.AddTextEffect( _
            PresetTextEffect:=msoTextEffect1, _
            Text:=classification, _
            FontName:="Arial", _
            FontSize:=60, _
            FontBold:=True, _
            FontItalic:=False, _
            Left:=100, _
            Top:=100)
        
        Selection.ShapeRange.TextEffect.NormalizedHeight = False
        Selection.ShapeRange.Line.Visible = False
        Selection.ShapeRange.Fill.Visible = True
        Selection.ShapeRange.Fill.Solid
        Selection.ShapeRange.Fill.ForeColor.RGB = watermarkColor
        Selection.ShapeRange.Fill.Transparency = 0.5
        Selection.ShapeRange.Rotation = 315
        Selection.ShapeRange.LockAspectRatio = True
        Selection.ShapeRange.Height = InchesToPoints(2.5)
        Selection.ShapeRange.Width = InchesToPoints(6)
        
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        
        MsgBox "Classification watermark has been added."
    Else
        MsgBox "No classification entered. Watermark not added."
    End If
End Sub
```

### 12. Create Security Header/Footer
Adds security classification information to headers and footers.

```vba
Sub CreateSecurityHeaderFooter()
    Dim classification As String
    classification = InputBox("Enter document classification (e.g., Public, Confidential, Secret):")
    
    If classification <> "" Then
        Dim section As Section
        
        For Each section In ActiveDocument.Sections
            ' Add header
            With section.Headers(wdHeaderFooterPrimary)
                .Range.Text = classification & " - " & Application.UserName & " - " & Format(Date, "yyyy-mm-dd")
                .Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            End With
            
            ' Add footer
            With section.Footers(wdHeaderFooterPrimary)
                .Range.Text = "Page " & section.Footers(wdHeaderFooterPrimary).PageNumbers.Add(PageNumberAlignment:=wdAlignPageNumberCenter)
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        Next section
        
        MsgBox "Security header and footer have been added."
    Else
        MsgBox "No classification entered. Header and footer not added."
    End If
End Sub
```

### 13. Create Security Incident Report Template
Creates a new security incident report template document.

```vba
Sub CreateSecurityIncidentReportTemplate()
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    With newDoc
        .Content.Text = "SECURITY INCIDENT REPORT" & vbCrLf & vbCrLf
        
        .Content.InsertAfter "Incident ID: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Date/Time Detected: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Reporter: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Incident Category: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Severity Level: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Description: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Impact Assessment: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Affected Systems: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Actions Taken: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Status: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Resolution: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Lessons Learned: " & vbCrLf & vbCrLf
        
        ' Format the title
        .Paragraphs(1).Range.Font.Bold = True
        .Paragraphs(1).Range.Font.Size = 16
        .Paragraphs(1).Alignment = wdAlignParagraphCenter
        
        ' Format the headings
        Dim i As Integer
        For i = 3 To .Paragraphs.Count Step 2
            If i < .Paragraphs.Count Then
                .Paragraphs(i).Range.Font.Bold = True
            End If
        Next i
        
        .SaveAs2 FileName:="Security_Incident_Report_Template.docx"
    End With
    
    MsgBox "Security Incident Report Template has been created."
End Sub
```

### 14. Create Security Assessment Template
Creates a new security assessment report template document.

```vba
Sub CreateSecurityAssessmentTemplate()
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    With newDoc
        .Content.Text = "SECURITY ASSESSMENT REPORT" & vbCrLf & vbCrLf
        
        .Content.InsertAfter "Assessment ID: " & vbCrLf & vbCrLf
        .Content.InsertAfter "System/Application Name: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Assessment Date: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Assessor: " & vbCrLf & vbCrLf
        .Content.InsertAfter "System Owner: " & vbCrLf & vbCrLf
        .Content.InsertAfter "System Classification: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Assessment Type: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Scope: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Methodology: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Executive Summary: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Findings:" & vbCrLf & vbCrLf
        .Content.InsertAfter "Risk Rating: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Recommendations: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Remediation Plan: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Conclusion: " & vbCrLf & vbCrLf
        .Content.InsertAfter "Attachments: " & vbCrLf & vbCrLf
        
        ' Format the title
        .Paragraphs(1).Range.Font.Bold = True
        .Paragraphs(1).Range.Font.Size = 16
        .Paragraphs(1).Alignment = wdAlignParagraphCenter
        
        ' Format the headings
        Dim i As Integer
        For i = 3 To .Paragraphs.Count Step 2
            If i < .Paragraphs.Count Then
                .Paragraphs(i).Range.Font.Bold = True
            End If
        Next i
        
        ' Add a table for findings
        Dim findingsTable As Table
        Set findingsTable = .Tables.Add(Range:=.Paragraphs(19).range, NumRows:=5, NumColumns:=4)
        
        With findingsTable
            .Cell(1, 1).Range.Text = "Finding ID"
            .Cell(1, 2).Range.Text = "Vulnerability"
            .Cell(1, 3).Range.Text = "Risk Level"
            .Cell(1, 4).Range.Text = "Recommendation"
            
            ' Format header row
            .Rows(1).Range.Font.Bold = True
            .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
        End With
        
        .SaveAs2 FileName:="Security_Assessment_Template.docx"
    End With
    
    MsgBox "Security Assessment Template has been created."
End Sub
```

### 15. Auto-number Security Findings
Automatically numbers paragraphs starting with "Finding".

```vba
Sub AutoNumberSecurityFindings()
    Dim range As range
    Dim findingCount As Integer
    
    findingCount = 0
    
    ' Look for paragraphs starting with "Finding"
    With ActiveDocument.Content.Find
        .Text = "Finding"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute
            Set range = Selection.range
            ' Check if it's at the beginning of a paragraph
            If range.Start = range.Paragraphs(1).range.Start Then
                findingCount = findingCount + 1
                ' Insert the finding number
                range.Collapse Direction:=wdCollapseStart
                range.Text = findingCount & ". "
                range.Collapse Direction:=wdCollapseEnd
                ' Move to end of paragraph to continue searching
                range.MoveEnd wdParagraph, 1
                Selection.Collapse wdCollapseEnd
                Selection.MoveTo range.End
            End If
        Loop
    End With
    
    MsgBox findingCount & " security findings have been numbered."
End Sub
```

### 16. Generate Table of Contents for Security Reports
Generates a table of contents for security reports.

```vba
Sub GenerateSecurityReportTOC()
    Dim tocRange As range
    Set tocRange = ActiveDocument.range(0, 0)
    
    ' Insert a page break at the beginning if there's content
    If ActiveDocument.Content.Text <> "" Then
        tocRange.Collapse Direction:=wdCollapseStart
        tocRange.InsertBreak Type:=wdPageBreak
        Set tocRange = ActiveDocument.range(0, 0)
    End If
    
    ' Add a title for the TOC
    tocRange.Text = "TABLE OF CONTENTS" & vbCrLf & vbCrLf
    tocRange.Style = wdStyleHeading1
    
    ' Add the TOC
    tocRange.Collapse Direction:=wdCollapseEnd
    ActiveDocument.TablesOfContents.Add range:=tocRange, _
        RightAlignPageNumbers:=True, _
        UseHeadingStyles:=True, _
        UpperHeadingLevel:=1, _
        LowerHeadingLevel:=3, _
        IncludePageNumbers:=True, _
        AddedStyles:="Security Finding"
    
    MsgBox "Table of Contents has been generated."
End Sub
```

## Document Analysis & Reporting

### 17. Generate Document Statistics Report
Creates a new document with statistics about the current document.

```vba
Sub GenerateDocumentStatisticsReport()
    Dim stats As String
    stats = "Document Statistics:" & vbCrLf & vbCrLf
    stats = stats & "Word Count: " & ActiveDocument.BuiltInDocumentProperties("Number of Words").Value & vbCrLf
    stats = stats & "Character Count: " & ActiveDocument.BuiltInDocumentProperties("Number of Characters").Value & vbCrLf
    stats = stats & "Paragraph Count: " & ActiveDocument.BuiltInDocumentProperties("Number of Paragraphs").Value & vbCrLf
    stats = stats & "Page Count: " & ActiveDocument.BuiltInDocumentProperties("Number of Pages").Value & vbCrLf
    stats = stats & "Creation Date: " & ActiveDocument.BuiltInDocumentProperties("Creation Date").Value & vbCrLf
    stats = stats & "Last Modified: " & ActiveDocument.BuiltInDocumentProperties("Last Save Time").Value & vbCrLf
    
    Dim newDoc As Document
    Set newDoc = Documents.Add
    newDoc.Content.Text = stats
    
    MsgBox "Document statistics report has been generated."
End Sub
```

### 18. Track Changes Summary
Generates a summary of tracked changes in the document.

```vba
Sub TrackChangesSummary()
    If Not ActiveDocument.TrackRevisions Then
        MsgBox "Track Changes is not enabled in this document."
        Exit Sub
    End If
    
    Dim revisionsCount As Long
    revisionsCount = ActiveDocument.Revisions.Count
    
    Dim summary As String
    summary = "Track Changes Summary:" & vbCrLf & vbCrLf
    summary = summary & "Total Revisions: " & revisionsCount & vbCrLf & vbCrLf
    
    Dim rev As Revision
    Dim insertCount As Long, deleteCount As Long, formatCount As Long
    
    For Each rev In ActiveDocument.Revisions
        Select Case rev.Type
            Case wdRevisionInsert
                insertCount = insertCount + 1
            Case wdRevisionDelete
                deleteCount = deleteCount + 1
            Case wdRevisionFormat
                formatCount = formatCount + 1
        End Select
    Next rev
    
    summary = summary & "Insertions: " & insertCount & vbCrLf
    summary = summary & "Deletions: " & deleteCount & vbCrLf
    summary = summary & "Formatting Changes: " & formatCount
    
    Dim newDoc As Document
    Set newDoc = Documents.Add
    newDoc.Content.Text = summary
    
    MsgBox "Track Changes summary has been generated."
End Sub
```

### 19. Extract Comments for Review
Extracts all comments to a new document for review.

```vba
Sub ExtractCommentsForReview()
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    Dim comment As Comment
    Dim count As Integer
    
    count = 0
    newDoc.Content.Text = "Extracted Comments for Review:" & vbCrLf & vbCrLf
    
    For Each comment In ActiveDocument.Comments
        count = count + 1
        newDoc.Content.InsertAfter "Comment #" & count & ":" & vbCrLf
        newDoc.Content.InsertAfter "Author: " & comment.Author & vbCrLf
        newDoc.Content.InsertAfter "Date: " & comment.Date & vbCrLf
        newDoc.Content.InsertAfter "Scope: " & comment.Scope & vbCrLf
        newDoc.Content.InsertAfter "Text: " & comment.Range.Text & vbCrLf
        newDoc.Content.InsertAfter "Comment: " & comment.Range.Comments(1).Range.Text & vbCrLf & vbCrLf
    Next comment
    
    If count = 0 Then
        newDoc.Content.InsertAfter "No comments found in the document."
    End If
    
    MsgBox count & " comments have been extracted for review."
End Sub
```

### 20. Check Document for Embedded Objects
Identifies and reports on embedded objects in the document.

```vba
Sub CheckForEmbeddedObjects()
    Dim newDoc As Document
    Set newDoc = Documents.Add
    
    Dim objectCount As Integer
    Dim inlineShape As InlineShape
    Dim shape As shape
    
    objectCount = 0
    newDoc.Content.Text = "Embedded Objects Report:" & vbCrLf & vbCrLf
    
    ' Check for inline shapes
    For Each inlineShape In ActiveDocument.InlineShapes
        objectCount = objectCount + 1
        newDoc.Content.InsertAfter "Object #" & objectCount & ":" & vbCrLf
        newDoc.Content.InsertAfter "Type: Inline Shape" & vbCrLf
        newDoc.Content.InsertAfter "Type Details: " & GetInlineShapeType(inlineShape.Type) & vbCrLf
        newDoc.Content.InsertAfter "Width: " & inlineShape.Width & vbCrLf
        newDoc.Content.InsertAfter "Height: " & inlineShape.Height & vbCrLf & vbCrLf
    Next inlineShape
    
    ' Check for shapes
    For Each shape In ActiveDocument.Shapes
        objectCount = objectCount + 1
        newDoc.Content.InsertAfter "Object #" & objectCount & ":" & vbCrLf
        newDoc.Content.InsertAfter "Type: Shape" & vbCrLf
        newDoc.Content.InsertAfter "Type Details: " & GetShapeType(shape.Type) & vbCrLf
        newDoc.Content.InsertAfter "Width: " & shape.Width & vbCrLf
        newDoc.Content.InsertAfter "Height: " & shape.Height & vbCrLf & vbCrLf
    Next shape
    
    If objectCount = 0 Then
        newDoc.Content.InsertAfter "No embedded objects found in the document."
    End If
    
    MsgBox objectCount & " embedded objects have been identified in the report."
End Sub

' Helper function to get inline shape type description
Function GetInlineShapeType(shapeType As WdInlineShapeType) As String
    Select Case shapeType
        Case wdInlineShapePicture
            GetInlineShapeType = "Picture"
        Case wdInlineShapeLinkedPicture
            GetInlineShapeType = "Linked Picture"
        Case wdInlineShapeOLEObject
            GetInlineShapeType = "OLE Object"
        Case wdInlineShapeLinkedOLEObject
            GetInlineShapeType = "Linked OLE Object"
        Case wdInlineShapeHorizontalLine
            GetInlineShapeType = "Horizontal Line"
        Case wdInlineShapeSmartArt
            GetInlineShapeType = "SmartArt"
        Case wdInlineShapeChart
            GetInlineShapeType = "Chart"
        Case wdInlineShapeLockedCanvas
            GetInlineShapeType = "Locked Canvas"
        Case Else
            GetInlineShapeType = "Unknown Type (" & shapeType & ")"
    End Select
End Function

' Helper function to get shape type description
Function GetShapeType(shapeType As MsoShapeType) As String
    Select Case shapeType
        Case msoAutoShape
            GetShapeType = "AutoShape"
        Case msoCallout
            GetShapeType = "Callout"
        Case msoCanvas
            GetShapeType = "Canvas"
        Case msoChart
            GetShapeType = "Chart"
        Case msoComment
            GetShapeType = "Comment"
        Case msoContentApp
            GetShapeType = "Content App"
        Case msoDiagram
            GetShapeType = "Diagram"
        Case msoEmbeddedOLEObject
            GetShapeType = "Embedded OLE Object"
        Case msoFormControl
            GetShapeType = "Form Control"
        Case msoFreeform
            GetShapeType = "Freeform"
        Case msoGroup
            GetShapeType = "Group"
        Case msoLine
            GetShapeType = "Line"
        Case msoLinkedOLEObject
            GetShapeType = "Linked OLE Object"
        Case msoLinkedPicture
            GetShapeType = "Linked Picture"
        Case msoMedia
            GetShapeType = "Media"
        Case msoOLEControlObject
            GetShapeType = "OLE Control Object"
        Case msoPicture
            GetShapeType = "Picture"
        Case msoPlaceholder
            GetShapeType = "Placeholder"
        Case msoScriptAnchor
            GetShapeType = "Script Anchor"
        Case msoShapeTypeMixed
            GetShapeType = "Mixed Type"
        Case msoTable
            GetShapeType = "Table"
        Case msoTextBox
            GetShapeType = "Text Box"
        Case msoTextEffect
            GetShapeType = "Text Effect"
        Case Else
            GetShapeType = "Unknown Type (" & shapeType & ")"
    End Select
End Function
```

### 21. Check Document for Accessibility Issues
Checks for common accessibility issues in the document.

```vba
Sub CheckAccessibilityIssues()
    Dim issues As String
    issues = "Accessibility Issues:" & vbCrLf & vbCrLf
    
    ' Check for missing alt text on images
    Dim shape As shape
    Dim missingAltText As Boolean
    
    missingAltText = False
    For Each shape In ActiveDocument.Shapes
        If shape.Type = msoPicture Or shape.Type = msoLinkedPicture Then
            If shape.Title = "" And shape.AlternativeText = "" Then
                missingAltText = True
                Exit For
            End If
        End If
    Next shape
    
    If missingAltText Then
        issues = issues & "- Some images are missing alternative text." & vbCrLf
    End If
    
    ' Check for proper heading structure
    Dim headingCount As Integer
    headingCount = 0
    
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 1" Then
            headingCount = headingCount + 1
        End If
    Next para
    
    If headingCount = 0 Then
        issues = issues & "- No Heading 1 styles found. Document structure may be unclear." & vbCrLf
    ElseIf headingCount > 1 Then
        issues = issues & "- Multiple Heading 1 styles found. Consider using only one Heading 1 for the document title." & vbCrLf
    End If
    
    ' Check for complex tables
    Dim tbl As Table
    Dim complexTable As Boolean
    
    complexTable = False
    For Each tbl In ActiveDocument.Tables
        If tbl.Columns.Count > 5 Or tbl.Rows.Count > 10 Then
            complexTable = True
            Exit For
        End If
    Next tbl
    
    If complexTable Then
        issues = issues & "- Complex tables detected. Ensure they have proper headers and descriptions." & vbCrLf
    End If
    
    ' Display results
    If issues = "Accessibility Issues:" & vbCrLf & vbCrLf Then
        issues = issues & "No obvious accessibility issues found."
    End If
    
    ' Create a new document with the results
    Dim newDoc As Document
    Set newDoc = Documents.Add
    newDoc.Content.Text = issues
    
    MsgBox "Accessibility check completed. Results have been generated in a new document."
End Sub
```

## Document Conversion & Export

### 22. Export Document to PDF with Security
Exports the document to PDF with security settings.

```vba
Sub ExportToPDFWithSecurity()
    Dim pdfPath As String
    pdfPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".")) & "pdf"
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False
    
    MsgBox "Document has been exported to PDF with security settings."
End Sub
```

### 23. Convert Document to Plain Text
Saves the document as a plain text file.

```vba
Sub ConvertToPlainText()
    Dim plainTextPath As String
    plainTextPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".")) & "txt"
    
    ' Save as plain text
    ActiveDocument.SaveAs2 FileName:=plainTextPath, FileFormat:=wdFormatText
    
    MsgBox "Document has been converted to plain text and saved as: " & plainTextPath
End Sub
```

### 24. Export Document Properties to CSV
Exports document properties to a CSV file.

```vba
Sub ExportDocumentPropertiesToCSV()
    Dim csvPath As String
    csvPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".")) & "csv"
    
    Dim fileNum As Integer
    fileNum = FreeFile()
    
    On Error Resume Next
    Open csvPath For Output As #fileNum
    If Err.Number <> 0 Then
        MsgBox "Error creating CSV file: " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Write CSV header
    Print #fileNum, "Property Name,Property Value"
    
    ' Write built-in properties
    Dim prop As DocumentProperty
    For Each prop In ActiveDocument.BuiltInDocumentProperties
        On Error Resume Next
        Print #fileNum, """" & prop.Name & """,""" & prop.Value & """"
        On Error GoTo 0
    Next prop
    
    ' Write custom properties
    For Each prop In ActiveDocument.CustomDocumentProperties
        On Error Resume Next
        Print #fileNum, """" & prop.Name & """,""" & prop.Value & """"
        On Error GoTo 0
    Next prop
    
    Close #fileNum
    
    MsgBox "Document properties have been exported to: " & csvPath
End Sub
```

## Document Management & Automation

### 25. Backup Current Document
Creates a timestamped backup of the current document.

```vba
Sub BackupCurrentDocument()
    Dim backupPath As String
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd_hh-mm-ss")
    backupPath = ActiveDocument.Path & "\Backup_" & timestamp & "_" & ActiveDocument.Name
    
    ActiveDocument.SaveAs FileName:=backupPath
    ActiveDocument.Save
    
    MsgBox "Document backup created at: " & backupPath
End Sub
```

### 26. Auto-save Document with Timestamp
Saves the document with a timestamp in the filename.

```vba
Sub AutoSaveWithTimestamp()
    Dim originalPath As String
    Dim timestamp As String
    Dim newPath As String
    
    originalPath = ActiveDocument.FullName
    timestamp = Format(Now(), "yyyy-mm-dd_hh-mm-ss")
    
    ' Create a new filename with timestamp
    newPath = Left(originalPath, InStrRev(originalPath, ".")) & timestamp & Mid(originalPath, InStrRev(originalPath, "."))
    
    ' Save the document with the new name
    ActiveDocument.SaveAs2 FileName:=newPath
    
    ' Reopen the original document
    Documents.Open FileName:=originalPath
    
    ' Close the timestamped version
    Documents(newPath).Close SaveChanges:=False
    
    MsgBox "Document has been auto-saved with timestamp: " & newPath
End Sub
```

### 27. Log Document Access
Logs document access to a text file.

```vba
Sub LogDocumentAccess()
    Dim logFilePath As String
    logFilePath = ActiveDocument.Path & "\DocumentAccessLog.txt"
    
    Dim logEntry As String
    logEntry = Now() & " - " & Application.UserName & " - " & Environ("COMPUTERNAME") & " - " & ActiveDocument.Name
    
    Dim fileNum As Integer
    fileNum = FreeFile()
    
    On Error Resume Next
    Open logFilePath For Append As #fileNum
    If Err.Number = 0 Then
        Print #fileNum, logEntry
        Close #fileNum
        MsgBox "Document access has been logged."
    Else
        MsgBox "Error logging document access: " & Err.Description
    End If
    On Error GoTo 0
End Sub
```

### 28. Batch Find and Replace Across Multiple Documents
Performs find and replace operations across multiple documents.

```vba
Sub BatchFindReplaceAcrossDocuments()
    Dim findText As String
    Dim replaceText As String
    
    findText = InputBox("Enter text to find:")
    If findText = "" Then Exit Sub
    
    replaceText = InputBox("Enter replacement text:")
    
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing Word documents"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Dim fileName As String
    Dim doc As Document
    Dim processedCount As Integer
    
    processedCount = 0
    fileName = Dir(folderPath & "\*.doc*")
    
    Application.ScreenUpdating = False
    
    Do While fileName <> ""
        Set doc = Documents.Open(FileName:=folderPath & "\" & fileName, Visible:=False)
        
        With doc.Content.Find
            .Text = findText
            .Replacement.Text = replaceText
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
        
        doc.Save
        doc.Close
        processedCount = processedCount + 1
        
        fileName = Dir()
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Find and replace completed in " & processedCount & " documents."
End Sub
```

### 29. Compare Document Versions
Compares two document versions and shows differences.

```vba
Sub CompareDocumentVersions()
    Dim originalPath As String, revisedPath As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Original Document"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.doc"
        If .Show = -1 Then
            originalPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Revised Document"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.doc"
        If .Show = -1 Then
            revisedPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.CompareDocuments OriginalDocument:=Documents.Open(originalPath), _
        RevisedDocument:=Documents.Open(revisedPath), _
        Destination:=wdCompareDestinationNew, _
        Granularity:=wdGranularityWordLevel
    
    MsgBox "Document comparison has been completed."
End Sub
```

### 30. Add Digital Signature
Adds a digital signature to the document.

```vba
Sub AddDigitalSignature()
    On Error Resume Next
    ActiveDocument.Signatures.Add
    If Err.Number <> 0 Then
        MsgBox "Error adding digital signature: " & Err.Description
    Else
        MsgBox "Digital signature has been added to the document."
    End If
    On Error GoTo 0
End Sub
```

## How to Use These Macros

1. Open Microsoft Word
2. Press `Alt + F11` to open the VBA Editor
3. In the VBA Editor, go to `Insert > Module` to create a new module
4. Copy and paste the desired macro code into the module
5. Close the VBA Editor
6. To run a macro:
   - Press `Alt + F8` to open the Macro dialog
   - Select the macro from the list
   - Click `Run`

## Contributing

Contributions are welcome! If you have additional VBA macros that would be useful for security information analysts, please submit a pull request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
