Sub ConvertFolderToPDF()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strFolderPath As String
    Dim strPDFPath As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder containing PowerPoint files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            strFolderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    Set pptApp = CreateObject("PowerPoint.Application")

    Set objFolder = CreateObject("Scripting.FileSystemObject").GetFolder(strFolderPath)
    For Each objFile In objFolder.Files
        If LCase(Right(objFile.Name, 4)) = ".ppt" Or LCase(Right(objFile.Name, 5)) = ".pptx" Then
            Set pptPresentation = pptApp.Presentations.Open(objFile.Path)
            strPDFPath = objFile.ParentFolder & "\" & Left(objFile.Name, InStrRev(objFile.Name, ".") - 1) & ".pdf"
            pptPresentation.SaveAs strPDFPath, 32 ' 32 denotes the format for PDF in PowerPoint
            pptPresentation.Close
        End If
    Next objFile
    pptApp.Quit

    Set pptPresentation = Nothing
    Set pptApp = Nothing
    Set objFile = Nothing
    Set objFolder = Nothing

    MsgBox "Conversion Complete", vbInformation
End Sub
