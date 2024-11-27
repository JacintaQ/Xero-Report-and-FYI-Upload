VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileSelectorForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1830
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "FileSelectorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonSelect_Click()
    ' Check if a file is selected
    If ComboBoxFiles.ListIndex = -1 Then
        MsgBox "Please select a file.", vbExclamation
    Else
        Dim selectedFile As String
        selectedFile = ComboBoxFiles.value
        MsgBox "Selected file: " & selectedFile, vbInformation
        ' Add further actions here, such as opening the file
    End If

    ' Close the UserForm
    Unload Me
End Sub



Private Sub UserForm_Initialize()
    Dim folderPath As String
    Dim folderName As String

    ' Set the main folder path
    folderPath = ActiveSheet.Range("B1").value & "\"

    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "The specified folder does not exist: " & folderPath, vbExclamation
        Unload Me
        Exit Sub
    End If

    ' Loop through each subfolder in the folder and add it to the ComboBox
    folderName = Dir(folderPath & "*", vbDirectory) ' Get the first entry (folder or file)

    Do While folderName <> ""
        ' Check if it's a folder and ignore "." and ".." entries
        If (GetAttr(folderPath & folderName) And vbDirectory) <> 0 Then
            If folderName <> "." And folderName <> ".." Then
                ComboBoxFiles.AddItem folderName ' Add the folder name to the ComboBox
            End If
        End If
        folderName = Dir ' Get the next entry
    Loop
End Sub

