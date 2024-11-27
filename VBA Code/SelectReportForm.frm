VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectReportForm 
   Caption         =   "Select Trial Balance Ending Year"
   ClientHeight    =   3060
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   6958
   OleObjectBlob   =   "SelectReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserCancel As Boolean

Private Sub LblClndr1_Click()
    ' Now uses ComboBoxYear1 only
    If ComboBoxYear1.value <> "" Then
        ComboBoxYear1.BackColor = RGB(255, 255, 255)
    Else
        MsgBox "Please select a valid year from the dropdown!", vbExclamation + vbOKOnly, "Invalid Year"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Clear and populate ComboBoxYear1 with years 2020-2024
    ComboBoxYear1.Clear
    Dim year As Integer
    For year = 2020 To 2025
        ComboBoxYear1.AddItem year
    Next year
    
    '' Other UI components
    'ComboBox1.Enabled = False
    'ComboBox1.value = "Profit & Loss Report"
    
    ' Command button set up
    cmdbCancel.Visible = True
    cmdbSubmit.Visible = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Me.hide
        UserCancel = True
    End If
End Sub

Private Sub cmdbCancel_Click()
    UserCancel = True
    Me.hide
End Sub

Private Sub cmdbSubmit_Click()
    Dim InvalidSubmit As Boolean
    InvalidSubmit = False
    
    Dim selectedYear As Integer
    
    'If ComboBox1.value = "" Then
    '    MsgBox "Please select a report type from the list!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    '    ComboBox1.SetFocus
    '    InvalidSubmit = True
    'End If
    
    If ComboBoxYear1.value = "" Then
        MsgBox "Please select the year of the reporting period!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ComboBoxYear1.BackColor = RGB(255, 230, 230)
        InvalidSubmit = True
    Else
        selectedYear = CInt(ComboBoxYear1.value)
    End If
    
    If Not InvalidSubmit Then
        'MsgBox "Trial balance report will be generated for the date ending in " & ComboBoxYear1.value & "-06-30"
        'MsgBox ComboBox1.value & " will be generated for the year " & _
            ComboBoxYear1.value, vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Me.hide
    End If
End Sub


