Attribute VB_Name = "Match_Files_FilePath"
' Upload folder name that needs to look up

Dim updateTime As Date
Dim combinedFolderPath As String ' Declare at the module level

' Start the auto-update and folder processing in one go
Sub StartProcess()
    UpdateSubfolders
    CombineFolderPaths
    Button1_Click
End Sub

' Update the folder

Sub UpdateSubfolders()
    ActiveSheet.Unprotect ""
    Dim folderPath As String
    Dim ws As Worksheet
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim subfolderList As String
    
    ' Set the worksheet
    Set ws = Sheets("File Path") '

    ' Get the base folder path from cell B1
    folderPath = ws.Range("B1").value ' Assuming K1 contains the main folder path

    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Invalid folder path. Please enter a valid path in cell B1."
        Exit Sub
    End If

    ' Set the folder object
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop through and list all subfolders
    For Each subfolder In folder.SubFolders
        subfolderList = subfolderList & subfolder.name & ","
    Next subfolder
    
    ActiveSheet.Protect ""
End Sub



Sub CombineFolderPaths()
    Dim baseFolderPath As String
    Dim subfolderName As String
    Dim ws As Worksheet
    
    ' Set the worksheet
    Set ws = Sheets("File Path") ' Change to your actual sheet name

    ' Get the base folder path from cell B1
    baseFolderPath = ws.Range("B1").value

    ' Check if base folder path is provided
    If baseFolderPath = "" Then
        MsgBox "Please enter the base folder path in cell B1."
        Exit Sub
    End If

    ' Get the selected subfolder name from cell K2
    subfolderName = ws.Range("A13").value

    ' Check if a subfolder has been selected
    If subfolderName = "" Then
        MsgBox "Please select a subfolder from the dropdown in cell A13."
        Exit Sub
    End If

    ' Combine the base folder path with the selected subfolder
    combinedFolderPath = baseFolderPath & "\" & subfolderName

    ' Display the combined folder path (optional)
    MsgBox "The combined folder path is: " & combinedFolderPath
End Sub

' Generate file and related filepath

Function FolderHasFiles(folderPath As String) As Boolean
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subfolder As Object
    Dim row As Long
    row = 17 ' Start row for displaying folder and file names

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder does not exist!"
        FolderHasFiles = False
        Exit Function
    End If
    
    ' Set folder object
    Set folder = fso.GetFolder(folderPath)
    
    ' Check if there are any files or subfolders
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        FolderHasFiles = False
        Exit Function
    End If

    ' List all files in the main folder
    For Each file In folder.Files
        Sheets("File Path").Range("H" & row).value = GetLastFolderName(folder.Path) ' Only show folder name in column H
            
        Sheets("File Path").Hyperlinks.Add _
            Anchor:=Sheets("File Path").Range("I" & row), _
            Address:=file.Path, _
            TextToDisplay:=file.name ' Add file name hyperlink in column I
        
        ' Display the file path in column J
        Sheets("File Path").Range("J" & row).value = file.Path ' Add file path in column J
        row = row + 1
    Next file

    ' Recursively list files in subfolders
    For Each subfolder In folder.SubFolders
        ListFilesInSubfolder subfolder, row
    Next subfolder

    FolderHasFiles = True
End Function

Sub ListFilesInSubfolder(subfolder As Object, ByRef row As Long)
    Dim file As Object
    Dim subsubfolder As Object

    ' List files in the current subfolder
    For Each file In subfolder.Files
        Sheets("File Path").Range("H" & row).value = GetLastFolderName(subfolder.Path) ' Only show folder name in column H
        
        ' Add file name hyperlink in column I
        Sheets("File Path").Hyperlinks.Add _
            Anchor:=Sheets("File Path").Range("I" & row), _
            Address:=file.Path, _
            TextToDisplay:=file.name
        
         ' Display the file path in column J
        Sheets("File Path").Range("J" & row).value = file.Path ' Add file path in column J
        row = row + 1
        
        ' Set column H width to J
        Sheets("File Path").columns("J").EntireColumn.Hidden = True
    Next file

    ' Recursively call for each subfolder
    For Each subsubfolder In subfolder.SubFolders
        ListFilesInSubfolder subsubfolder, row
    Next subsubfolder
End Sub

Function GetLastFolderName(folderPath As String) As String
    ' Split the folder path by backslash and return the last part
    Dim parts As Variant
    parts = Split(folderPath, "\")
    GetLastFolderName = parts(UBound(parts))
End Function



Sub Button1_Click()
    ActiveSheet.Unprotect ""
    ' Clear data in columns H and I
    Sheets("File Path").Range("H17:J" & Sheets("File Path").Rows.Count).ClearContents

    ' Display clickable links of file names in column F
    If FolderHasFiles(combinedFolderPath) Then
        ' MsgBox "Files exist in the folder!"
    Else
        MsgBox "No files found in the folder!"
        Exit Sub
    End If
    ActiveSheet.Protect ""
End Sub


Sub UpdateLinksInColumnE()
    ActiveSheet.Unprotect ""
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim startTime As Double
    Dim elapsedTime As Double
    Dim estimatedTimeLeft As Double
    Dim minutesLeft As Long
    Dim secondsLeft As Long

    ' Define worksheet
    Set ws = ThisWorkbook.Sheets("File Path") ' Adjust worksheet name if necessary

    ' Clear data in columns E and F
    ws.Range("E17:E" & ws.Rows.Count).ClearContents
    ws.Range("G2").ClearContents
    
    ' Find the last row in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Start the timer
    startTime = Timer

    ' Loop through each row starting from row 8
    For i = 17 To lastRow
    
            ' Check if column B is empty
        If ws.Cells(i, "A").value = "" And ws.Cells(i, "B").value = "" And ws.Cells(i, "C").value = "" And ws.Cells(i, "D").value = "" Then
            Exit For ' Break the loop if column B is empty
        End If

        ' Check if Column D is empty
        If ws.Cells(i, "D").value = "" Then
            ' If Column D is empty, clear Column E
            ws.Cells(i, "E").value = ""
        Else
            matchFound = False
            
            ' Loop through Column I to check if it contains Column A value
            For j = 8 To lastRow
                If InStr(1, ws.Cells(j, "I").value, ws.Cells(i, "A").value) > 0 Then
                    ' Add hyperlink in Column E and set background to light blue
                    ws.Hyperlinks.Add _
                        Anchor:=ws.Cells(i, "E"), _
                        Address:=ws.Cells(j, "J").value, _
                        TextToDisplay:=ws.Cells(j, "I").value
                    ws.Cells(i, "E").Font.Bold = True
                    matchFound = True
                    Exit For
                End If
            Next j
            
            ' If no match found with Column A, check if Column H contains Column D value
            If Not matchFound Then
                For j = 17 To lastRow
                    If InStr(1, ws.Cells(j, "I").value, ws.Cells(i, "D").value) > 0 Then
                        ' Add hyperlink in Column E and set background to light green
                        ws.Hyperlinks.Add _
                            Anchor:=ws.Cells(i, "E"), _
                            Address:=ws.Cells(j, "J").value, _
                            TextToDisplay:=ws.Cells(j, "I").value
                        
                        matchFound = True
                        Exit For
                    End If
                Next j
            End If
            
            ' If no match found, leave Column E empty
            If Not matchFound Then
                ws.Cells(i, "E").value = ""
                ws.Cells(i, "E").Interior.ColorIndex = xlNone ' Clear any previous color
            End If
        End If
        
        ' Update estimated time left in G2 every 50 rows
        If (i - 7) Mod 20 = 0 Then
            elapsedTime = Timer - startTime
            estimatedTimeLeft = (elapsedTime / (i - 7)) * (lastRow - i)
            
            ' Convert to minutes and seconds format
            minutesLeft = Int(estimatedTimeLeft / 60)
            secondsLeft = Int(estimatedTimeLeft Mod 60)
            
            ws.Cells(6, "D").value = "Estimated time left: " & minutesLeft & ":" & Format(secondsLeft, "00") & " min"
        End If
    Next i

    ' Final message when complete
    ws.Cells(6, "D").value = "Processing complete"
    ActiveSheet.Protect ""
End Sub


' Hide and Unhide Column H,I and K

Sub ToggleColumnsGHJ()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Use the currently active worksheet
    ActiveSheet.Unprotect ""
    ' Check if columns H, I, and K are hidden
    If ws.columns("H:I").EntireColumn.Hidden = True And ws.columns("K").EntireColumn.Hidden = True Then
        ' If hidden, unhide these columns
        ws.columns("H:I").EntireColumn.Hidden = False
        'ws.columns("J").EntireColumn.Hidden = False
        ws.columns("K").EntireColumn.Hidden = False
    Else
        ' If visible, hide these columns
        ws.columns("H:I").EntireColumn.Hidden = True
        'ws.columns("J").EntireColumn.Hidden = True
        ws.columns("K").EntireColumn.Hidden = True
    End If
    ActiveSheet.Protect ""
End Sub





