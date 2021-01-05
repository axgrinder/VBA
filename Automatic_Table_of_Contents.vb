Option Explicit
'Cell A1 on each page is the title for the TOC
Sub Auto_Table_Contents()

    Dim StartCell As Range 'for inputbox to select range
    Dim Sh As Worksheet
    Dim ShName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult 'for message box confirmation
    
    MsgConfirm = VBA.MsgBox("the Values in the cells could be overwritten. Would you like to continue?", vbOKCancel + vbDefaultButton2)
    If MsgConfirm = vbCancel Then Exit Sub
    
    On Error Resume Next
    
    Set StartCell = Excel.Application.InputBox("Where do you want to insert the table of contents?" _
    & vbNewLine & "Please Select a Cell:", "Insert Table of Contents", , , , , , 8)
    
    If Err.Number = 424 Then Exit Sub
    On Error GoTo handle
    
    Set StartCell = StartCell.Cells(1, 1)
    
    For Each Sh In Worksheets
        ShName = Sh.Name
        If ActiveSheet.Name <> ShName Then
            If Sh.Visible = xlSheetVisible Then
                ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _
                "'" & ShName & "'!A1", TextToDisplay:=ShName
                StartCell.Offset(0, 1).Value = Sh.Range("A1").Value
                Set StartCell = StartCell.Offset(1, 0)
            End If 'visible?
        End If 'sheet not active
    Next Sh
    Exit Sub
handle:
MsgBox "Unfortunately, and error has occurred."

End Sub
