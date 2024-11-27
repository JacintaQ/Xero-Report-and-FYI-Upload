Attribute VB_Name = "FYILinks_WKP"
Sub CopyFYILinksToOffsetColumnB(accountName As String, targetCell As Range, displayColumn As String)
    Dim wsSource As Worksheet
    Dim linkCell As Range
    Dim foundCell As Range
    Dim displayText As String
    
    ' Unprotect the target sheet
    targetCell.Worksheet.Unprotect ""
    
    On Error GoTo ProtectSheet ' Error handling to ensure protection in case of errors
    
    ' Set the source worksheet to "CheckFileExists"
    Set wsSource = ThisWorkbook.Sheets("File Path")
    
    ' Check if the account name is valid
    If accountName <> "" And accountName <> "0" Then
        Debug.Print accountName
        ' Search for the account name in column B of "CheckFileExists"
        Set foundCell = wsSource.columns("B").Find(What:=accountName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' If the account is found
        If Not foundCell Is Nothing Then
            ' Get the FYI link from the corresponding cell in column F
            Set linkCell = wsSource.Cells(foundCell.row, "F")
            displayText = wsSource.Cells(foundCell.row, displayColumn).value
            
            ' Clear any existing hyperlink in the target cell
            targetCell.ClearContents
            targetCell.ClearFormats
            If targetCell.Hyperlinks.Count > 0 Then
                targetCell.Hyperlinks.Delete
            End If
            
            ' Add the hyperlink to the target cell with display text
            If linkCell.value <> "" Then
                targetCell.Worksheet.Hyperlinks.Add Anchor:=targetCell, _
                                                    Address:=linkCell.value, _
                                                    TextToDisplay:=displayText
            End If
            
            ' Apply formatting to the target cell
            With targetCell
                .Interior.Color = RGB(217, 217, 217) ' Light gray background
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .EntireColumn.AutoFit
                .EntireRow.AutoFit
                .Locked = False ' Set the cell to be unlocked
            End With
        End If
    End If

ProtectSheet:
    ' Protect the sheet again
    targetCell.Worksheet.Protect ""
End Sub


' Clear the cell if no account is chosen

Sub ClearHyperlinkAndFormat(accountName As String, targetCell As Range)
    ' Remove hyperlink and contents if accountName is blank
    ActiveSheet.Unprotect ""
    If accountName = "" Then
        If targetCell.Hyperlinks.Count > 0 Then
            targetCell.Hyperlinks.Delete
        End If
        targetCell.ClearContents
    Else
        ' If accountName is not blank, add or update the hyperlink
        'Call AddOrUpdateHyperlink(accountName, targetCell)
    End If

    ' Apply formatting to the target cell after the hyperlink update
    With targetCell
        .Interior.Color = RGB(217, 217, 217) ' Light gray background
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .Locked = False ' Set the cell to be unlocked

    End With
    targetCell.Worksheet.Protect ""
End Sub

