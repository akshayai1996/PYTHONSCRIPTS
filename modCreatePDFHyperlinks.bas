' ===========================================================================
' Module Name  : modCreatePDFHyperlinks
' Description  :
'   This module provides macros to create clickable PDF hyperlinks in Excel
'   for selected cells containing file names (without extensions).
'
'   Main Features:
'     - Creates hyperlinks pointing to <ServerFolder>\<CellText>.pdf
'     - Keeps the visible cell text exactly as entered (no .pdf shown)
'     - Hyperlinks do NOT turn purple after clicking
'     - Clears cells if the target PDF does not exist to prevent broken links
'     - Works with shared files via Teams or locally downloaded copies
'     - Optional add-on macro to create hyperlinks only for newly added files
'
' Usage Instructions:
'   1. Select the range of cells containing the PDF file names.
'   2. Run `CreatePDFHyperlinks_NoPurple` to create hyperlinks for all selected cells.
'   3. To add hyperlinks for newly added files only, select the relevant cells and
'      run `AddHyperlinksToNewFiles`.
'   4. If a PDF does not exist at the given location, the cell is cleared.
'   5. During execution, the macro will prompt to select the base server folder
'      containing the PDFs.
'
' Example:
'   Cell value: "Valve_123"
'   Base folder selected: "\\Server\Projects\FRI_PDFs\"
'   Resulting hyperlink points to: "\\Server\Projects\FRI_PDFs\Valve_123.pdf"
'   Visible cell text remains: "Valve_123"
'
' Author       : Akshay Solanki
' Version      : 1.0
' Created On   : 19-Oct-2025
' Dependencies : None (built-in Excel VBA only)
' ===========================================================================

Attribute VB_Name = "modCreatePDFHyperlinks"
Option Explicit

' ================================
' Macro: Create hyperlinks for selected cells
' ================================
Public Sub CreatePDFHyperlinks_NoPurple()
    Dim cell As Range
    Dim baseFolder As String
    Dim created As Long, cleared As Long, processed As Long
    Dim nameOnly As String, fullPath As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells containing file names (without extension).", vbExclamation, "No range selected"
        Exit Sub
    End If

    ' Ask for a base server folder once
    baseFolder = PickBaseFolder("")
    If Len(baseFolder) = 0 Then
        MsgBox "No folder selected. Operation cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    
    If Len(baseFolder) > 0 Then baseFolder = EnsureTrailingSep(baseFolder)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    For Each cell In Selection.Cells
        processed = processed + 1
        nameOnly = Trim$(CStr(cell.Value))

        ' Remove any existing hyperlink so we start clean
        On Error Resume Next
        If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete
        On Error GoTo SafeExit

        ' If cell is empty, skip it
        If Len(nameOnly) = 0 Then GoTo NextCell

        ' Build target path:
        If IsLikelyFullPath(nameOnly) Then
            If HasExtension(nameOnly) Then
                fullPath = nameOnly
            Else
                fullPath = nameOnly & ".pdf"
            End If
        ElseIf Len(baseFolder) > 0 Then
            If HasExtension(nameOnly) Then
                fullPath = baseFolder & nameOnly
            Else
                fullPath = baseFolder & nameOnly & ".pdf"
            End If
        Else
            fullPath = vbNullString
        End If

        ' If we have a path, check existence; else clear cell
        If Len(fullPath) > 0 Then
            If Dir(fullPath, vbNormal) <> vbNullString Then
                On Error Resume Next
                cell.Hyperlinks.Add Anchor:=cell, Address:=fullPath, SubAddress:="", TextToDisplay:=nameOnly
                On Error GoTo SafeExit
                created = created + 1
            Else
                cell.Value = ""
                cell.ClearFormats
                cleared = cleared + 1
            End If
        Else
            cell.Value = ""
            cell.ClearFormats
            cleared = cleared + 1
        End If

NextCell:
    Next cell

    ' Align Followed Hyperlink style color so it does not turn purple
    AlignFollowedHyperlinkStyle True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Hyperlinks created: " & created & vbCrLf & _
           "Cells cleared (file not found): " & cleared & vbCrLf & _
           "Total processed: " & processed, _
           vbInformation, "PDF Hyperlinking Completed"
End Sub

' ================================
' Macro: Add hyperlinks to newly added files
' ================================
Public Sub AddHyperlinksToNewFiles()
    Dim cell As Range
    Dim baseFolder As String
    Dim created As Long, skipped As Long, processed As Long
    Dim nameOnly As String, fullPath As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells to add hyperlinks to.", vbExclamation, "No range selected"
        Exit Sub
    End If

    baseFolder = PickBaseFolder("")
    If Len(baseFolder) = 0 Then
        MsgBox "No folder selected. Operation cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    
    If Len(baseFolder) > 0 Then baseFolder = EnsureTrailingSep(baseFolder)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit2

    For Each cell In Selection.Cells
        processed = processed + 1
        nameOnly = Trim$(CStr(cell.Value))

        If Len(nameOnly) = 0 Then GoTo NextCell2

        On Error Resume Next
        If cell.Hyperlinks.Count > 0 Then
            skipped = skipped + 1
            GoTo NextCell2
        End If
        On Error GoTo SafeExit2

        If IsLikelyFullPath(nameOnly) Then
            If HasExtension(nameOnly) Then
                fullPath = nameOnly
            Else
                fullPath = nameOnly & ".pdf"
            End If
        ElseIf Len(baseFolder) > 0 Then
            If HasExtension(nameOnly) Then
                fullPath = baseFolder & nameOnly
            Else
                fullPath = baseFolder & nameOnly & ".pdf"
            End If
        Else
            fullPath = vbNullString
        End If

        If Len(fullPath) > 0 Then
            If Dir(fullPath, vbNormal) <> vbNullString Then
                On Error Resume Next
                cell.Hyperlinks.Add Anchor:=cell, Address:=fullPath, SubAddress:="", TextToDisplay:=nameOnly
                On Error GoTo SafeExit2
                created = created + 1
            End If
        End If

NextCell2:
    Next cell

    AlignFollowedHyperlinkStyle True

SafeExit2:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "New hyperlinks created: " & created & vbCrLf & _
           "Cells skipped (already have hyperlinks): " & skipped & vbCrLf & _
           "Total processed: " & processed, _
           vbInformation, "New Files Hyperlinking Completed"
End Sub

' ================================
' Helper Functions
' ================================
Private Sub AlignFollowedHyperlinkStyle(Optional ByVal matchUnderline As Boolean = True)
    Dim wb As Workbook
    Dim stHyp As Style, stFol As Style
    On Error Resume Next
    Set wb = ActiveWorkbook
    Set stHyp = wb.Styles("Hyperlink")
    Set stFol = wb.Styles("Followed Hyperlink")
    If Not stHyp Is Nothing And Not stFol Is Nothing Then
        stFol.Font.Color = stHyp.Font.Color
        If matchUnderline Then stFol.Font.Underline = stHyp.Font.Underline
    End If
    On Error GoTo 0
End Sub

Private Function EnsureTrailingSep(ByVal path As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    If Len(path) = 0 Then
        EnsureTrailingSep = path
    ElseIf Right$(path, 1) = sep Then
        EnsureTrailingSep = path
    Else
        EnsureTrailingSep = path & sep
    End If
End Function

Private Function HasExtension(ByVal f As String) As Boolean
    Dim dotPos As Long
    dotPos = InStrRev(f, ".")
    HasExtension = (dotPos > 1 And dotPos < Len(f))
End Function

Private Function IsLikelyFullPath(ByVal p As String) As Boolean
    If Len(p) >= 2 Then
        IsLikelyFullPath = (Left$(p, 2) = "\\") Or _
                           (Mid$(p, 2, 2) = ":\") Or (Mid$(p, 2, 2) = ":/")
    Else
        IsLikelyFullPath = False
    End If
End Function

Private Function PickBaseFolder(ByVal defaultPath As String) As String
    Dim fd As FileDialog, showIt As Long
    Dim startPath As String

    startPath = defaultPath
    If Len(startPath) > 0 Then
        If Dir(startPath, vbDirectory) = vbNullString Then
            startPath = vbNullString
        End If
    End If

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Base Server Folder for PDF files"
        If Len(startPath) > 0 Then .InitialFileName = startPath
        showIt = .Show
        If showIt = -1 Then
            PickBaseFolder = EnsureTrailingSep(.SelectedItems(1))
        Else
            PickBaseFolder = vbNullString
        End If
    End With
End Function
