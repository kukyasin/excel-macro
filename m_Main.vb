Option Explicit

Public currentDoc As Document

Sub Main()

Application.ScreenUpdating = False

Dim inFolder As String, outFolder As String
Dim imagePath As String

Dim inFile As String, outFile As String
Dim objFolder As Object

Dim objFSO As Object
Dim objFile As Object

' Prompt the user for an action
Dim action As Integer
action = MsgBox("Choose an action: " & vbCrLf & "1: Search for a text string" & vbCrLf & _
    "2: Search and replace a text string" & vbCrLf & _
    "3: Replace logo", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Choose action")

' Exit if no action is chosen
If action = vbCancel Then
    Debug.Print "No action chosen."
    Exit Sub
End If

' Prompt the user for input
Dim oldText As String, newText As String
Dim newClassification As String

If action = vbYes Then ' 1: Search for a text string
    oldText = InputBox("Enter the text to find:", "Find Text")
    If oldText = "" Then
        Debug.Print "Nothing to search for."
        Exit Sub
    End If
ElseIf action = vbNo Then ' 2: Search and replace a text string
    oldText = InputBox("Enter the text to find:", "Find Text")
    If oldText = "" Then
        Debug.Print "Nothing to replace."
        Exit Sub
    End If
    newText = InputBox("Enter the New Text:", "Replace With")
    ' Get Output Folder
    Set objFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With objFolder
        .Title = "Select OUPUT Folder"
        .AllowMultiSelect = False
        If .Show = True Then
            outFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    If Right(outFolder, 1) <> "\" Then
        outFolder = outFolder & "\"
    End If
Else ' 3: Replace logo
    ' Get Image File
    imagePath = getimagePath()
End If

' Get Input Folder
Set objFolder = Application.FileDialog(msoFileDialogFolderPicker)
With objFolder
    .Title = "Select INPUT Folder"
    .AllowMultiSelect = False
    If .Show = True Then
        inFolder = .SelectedItems(1)
    Else
        Exit Sub
    End If
End With

If Right(inFolder, 1) <> "\" Then
    inFolder = inFolder & "\"
End If

' Loop through all files
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(inFolder)

For Each objFile In objFolder.Files
    inFile = objFile.Name
    If Right(inFile, 5) = ".docx" Then 'Specify the file extension you want to open
        ' Open File
        Set currentDoc = Documents.Open(objFile.Path)
        Application.ScreenUpdating = True
        ThisDocument.ContentControls(1).Range.Text = "Documents processed."
        Application.ScreenUpdating = False

        ' Process File
        
        ' Get new Name
        outFile = Replace(inFile, ".docx", " V1 DRAFT 00.docx")
        
        ' Process Document
        ' Search or Find and Replace
        If action = vbYes Then
            Call SearchText(oldText)
        ElseIf action = vbNo Then
            Call ReplaceText(oldText, newText)
        End If
        
        If action = vbRetry Then
            replaceImage (imagePath)
        End If

        
        ' Save As new name
        If action = vbNo Then
            currentDoc.SaveAs2 FileName:=outFolder & outFile, FileFormat:=wdFormatDocumentDefault
        End If
        
        ' Close old file
        currentDoc.Close SaveChanges:=False 'Close the document without saving changes
    End If

' go to next
Next objFile

' inform user
Application.ScreenUpdating = True
ThisDocument.ContentControls(1).Range.Text = "Documents processed."

End Sub
