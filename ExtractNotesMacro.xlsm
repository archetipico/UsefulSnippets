' Extract notes from Excel spreadsheet selections
Sub ExtractAnnotations()
    Dim cell As Range
    Dim annotations As String
    Dim clipboard As Object
    
    ' Initialize the annotations string
    annotations = ""
    
    ' Loop through each selected cell
    For Each cell In Selection
        ' Check if the cell has a comment
        If Not cell.Comment Is Nothing Then
            ' Add the annotation to the annotations string
            annotations = annotations & cell.Address & ": " & cell.Comment.Text & vbCrLf
        End If
    Next cell
    
    ' Copy the annotations to the clipboard
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.SetText annotations
    clipboard.PutInClipboard
End Sub
