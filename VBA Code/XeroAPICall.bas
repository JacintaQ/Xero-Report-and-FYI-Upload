Attribute VB_Name = "XeroAPICall"

' XeroAPICall v1.0.0
' @author musayohanes00@gmail.com
' https://github.com/Muyoouu/vba-xero-api
'
' Xero accounting API calls
' Docs: https://developer.xero.com/documentation/api/accounting/overview

Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' Provide the Xero client ID and client secret through these constants.
' Leave these constants empty to be prompted for the values during runtime.
Private Const cXEROCLIENTID As String = ""
Private Const cXEROCLIENTSECRET As String = ""

' Prefix used for naming the output sheet where the Profit and Loss report will be generated.
Private Const ReportOutputSheet As String = "TrialBalance_Report_"

' Used for naming the JSON file output, if any.
Private pLastOutputSheetName As String

' WebClient instance used for making API calls to Xero.
Private pXeroClient As WebClient

' Xero client ID and client secret values used for authentication.
Private pXeroClientId As String
Private pXeroClientSecret As String
Public LastSelectedTenantId As String
' --------------------------------------------- '
' Private Properties and Methods
' --------------------------------------------- '

''
' Retrieves the Xero API client ID.
' If the client ID is not provided through the 'cXEROCLIENTID' constant, the user is prompted to enter the client ID.
'
' @property XeroClientId
' @type {String}
' @return {String} The Xero API client ID.
''
Private Property Get XeroClientId() As String
    If pXeroClientId = "" Then
        If cXEROCLIENTID <> "" Then
            pXeroClientId = cXEROCLIENTID
        Else
            Dim InpBxResponse As String
            InpBxResponse = InputBox("Please Enter Xero API Client ID", "Xero Report Generator - Microsoft Excel")
            If InpBxResponse <> "" Then
                pXeroClientId = InpBxResponse
            Else
                Err.Raise 11041 + vbObjectError, "XeroAPICall.ClientIdInputBox", "User did not provide Xero API Client ID"
            End If
        End If
    End If
    
    XeroClientId = pXeroClientId
End Property

''
' Retrieves the Xero API client secret.
' If the client secret is not provided through the 'cXEROCLIENTSECRET' constant, the user is prompted to enter the client secret.
'
' @property XeroClientSecret
' @type {String}
' @return {String} The Xero API client secret.
''
Private Property Get XeroClientSecret() As String
    If pXeroClientSecret = "" Then
        If cXEROCLIENTSECRET <> "" Then
            pXeroClientSecret = cXEROCLIENTSECRET
        Else
            Dim InpBxResponse As String
            InpBxResponse = InputBox("Please Enter Xero API Client Secret", "Xero Report Generator - Microsoft Excel")
            If InpBxResponse <> "" Then
                pXeroClientSecret = InpBxResponse
            Else
                Err.Raise 11041 + vbObjectError, "XeroAPICall.ClientSecretInputBox", "User did not provide Xero API Client Secret"
            End If
        End If
    End If
    
    XeroClientSecret = pXeroClientSecret
End Property

''
' Initializes and returns a WebClient instance configured for making API calls to Xero.
'
' @property XeroClient
' @type {WebClient}
' @return {WebClient} The configured WebClient instance.
'
' The WebClient instance is set up with the following configurations:
' - Base URL set to 'https://api.xero.com/'
' - Authenticator set to an instance of the 'XeroAuthenticator' class, which handles Xero's OAuth2 authentication flow.
' - The 'offline_access' and 'accounting.reports.read' scopes are requested during the authentication process.
'
' The WebClient instance is cached and reused between requests.
''
Private Property Get XeroClient() As WebClient
    If pXeroClient Is Nothing Then
        ' Create a new WebClient instance with the base URL
        Set pXeroClient = New WebClient
        pXeroClient.BaseUrl = "https://api.xero.com/"
        
        ' Set up the 'XeroAuthenticator' instance for OAuth2 authentication
        Dim Auth As XeroAuthenticator
        Set Auth = New XeroAuthenticator
        Auth.Setup CStr(XeroClientId), CStr(XeroClientSecret)
        
        ' Request the 'offline_access' and 'accounting.reports.read' scopes
        Auth.AddScope "offline_access"
        Auth.AddScope "accounting.settings.read"
        Auth.AddScope "accounting.reports.read"
        
        ' Set the 'XeroAuthenticator' instance as the authenticator for the WebClient
        Set pXeroClient.Authenticator = Auth
        
    End If
    
    Set XeroClient = pXeroClient
End Property

''
' Sets the WebClient instance used for making API calls to Xero.
'
' @property XeroClient
' @type {WebClient}
' @param {WebClient} Client - The WebClient instance to set.
''
Private Property Set XeroClient(Client As WebClient)
    Set pXeroClient = Client
    
End Property

' Displays a user form that allows the user to select the report parameters (date range) for the Xero API request.
'
' @method SelectReport
' @param {WebRequest} Request - The WebRequest object to which the selected report parameters will be added.
' @return {WebRequest} The WebRequest object with the selected report parameters added as query string parameters.
'
' This function performs the following steps:
' 1. Initializes and displays the 'SelectReportForm' user form.
' 2. If the user cancels the form, raises an error and displays a message.
' 3. Converts the selected date range from the user form to the required format for the Xero API request.
' 4. Adds the 'fromDate' and 'toDate' query string parameters to the WebRequest object with the selected date range.
' 5. Returns the updated WebRequest object.
'
' Note: This function uses the 'TextBox1' and 'TextBox2' controls of the 'SelectReportForm' user form to retrieve the selected date range.

Private Function SelectReport(Request As WebRequest) As WebRequest
    On Error GoTo ApiCall_Cleanup

    ' Initialize the 'SelectReportForm' user form
    Dim SelectForm1 As SelectReportForm
    Set SelectForm1 = New SelectReportForm
    
    ' Display the user form
    SelectForm1.show
    
    ' Check if the user canceled the form
    If SelectForm1.UserCancel Then
        ' Notify the user and raise an error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly
        Err.Raise 11040 + vbObjectError, "SelectReportForm", "User canceled selection form"
    End If
    
    ' Convert the selected date range to the required format
    Dim fromDate As Date
    ' Dim toDate As Date
    
    fromDate = SelectForm1.ComboBoxYear1.value 'Revise to year only
    ' fromDate = DateSerial(CInt(Right(SelectForm1.TextBox1.value, 4)), CInt(Left(SelectForm1.TextBox1.value, 2)), CInt(Mid(SelectForm1.TextBox1.value, 4, 2)))
    ' toDate = DateSerial(CInt(Right(SelectForm1.TextBox2.value, 4)), CInt(Left(SelectForm1.TextBox2.value, 2)), CInt(Mid(SelectForm1.TextBox2.value, 4, 2)))
    
    ' Add the 'Date' and 'toDate' query string parameters to the WebRequest object
    Request.AddQuerystringParam "Date", Format(DateSerial(fromDate, 6, 30), "yyyy-mm-dd")
    ' Request.AddQuerystringParam "toDate", Format(toDate, "yyyy-mm-dd")
    
    ' Return the updated WebRequest object
    Set SelectReport = Request

ApiCall_Cleanup:
    ' Unload the user form and handle errors
    If Not SelectForm1 Is Nothing Then
        Unload SelectForm1
    End If
End Function



''
' Retrieves a TrialBalance report from the Xero API for the selected date range. #Revise
'
' @method GetPnLReport
' @return {Dictionary} A dictionary containing the Profit and Loss report data, or an empty dictionary if an error occurs.
'
' This function performs the following steps:
' 1. Initializes a new WebRequest object for the API request.
' 2. Configures the WebRequest object with the required parameters for the Profit and Loss report API endpoint.
' 3. Displays the 'SelectReportForm' user form to allow the user to select the report date range.
' 4. Sends the API request to the Xero API using the configured WebRequest object.
' 5. If the API request is successful (200 status code), returns the report data as a dictionary.
' 6. If the API request fails, raises an error with the appropriate error details.
'
' Note: This function uses the 'XeroClient' property to execute the API request and the 'SelectReport' function to obtain the report date range.
''
Private Function GetChartOfAccounts() As Dictionary
    On Error GoTo ApiCall_Cleanup

    
    ' Use the organization for the API request to get the chart of accounts
    Dim AccountsRequest As WebRequest
    Set AccountsRequest = New WebRequest
    AccountsRequest.Resource = "api.xro/2.0/Accounts"
    AccountsRequest.Method = WebMethod.HttpGet
    AccountsRequest.RequestFormat = WebFormat.FormUrlEncoded
    AccountsRequest.ResponseFormat = WebFormat.Json
    
    ' Send the API request and retrieve the response
    Dim AccountsResponse As WebResponse
    Set AccountsResponse = XeroClient.Execute(AccountsRequest)
    

    ' Initialize the 'SelectReportForm' user form
    Dim SelectForm1 As SelectReportForm
    Set SelectForm1 = New SelectReportForm
    
    ' Display the user form
    SelectForm1.show
    
    ' Check if the user canceled the form
    If SelectForm1.UserCancel Then
        ' Notify the user and raise an error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly
        Err.Raise 11040 + vbObjectError, "SelectReportForm", "User canceled selection form"
    End If

    ' Retrieve the selected values from the form after it's hidden
    Dim selectedYear As Integer
    Dim previousYear As Integer
    
    
    selectedYear = CInt(SelectForm1.ComboBoxYear1.value)
    previousYear = selectedYear - 1

'    ' Include the box that enable Select year
'    Dim selectedYear As Integer
'    Dim previousYear As Integer
'    selectedYear = CInt(InputBox("Please input Ending Year. e.g.:2024", "Trial Balance Report"))
'    previousYear = selectedYear - 1
    
    ' Initialize a new WebRequest object for the API request for the CurrentYear Trial Balance
    Dim ReportRequestCurrentYear As WebRequest
    Set ReportRequestCurrentYear = New WebRequest
    'ReportRequestCurrentYear.Resource = "api.xro/2.0/Reports/TrialBalance?Date=CurrentYear-06-30"
    ReportRequestCurrentYear.Resource = "api.xro/2.0/Reports/TrialBalance?Date=" & selectedYear & "-06-30"
    ReportRequestCurrentYear.Method = WebMethod.HttpGet
    ReportRequestCurrentYear.RequestFormat = WebFormat.FormUrlEncoded
    ReportRequestCurrentYear.ResponseFormat = WebFormat.Json

    ' Send the API request for CurrentYear data and retrieve the response
    Dim ReportResponseCurrentYear As WebResponse
    Set ReportResponseCurrentYear = XeroClient.Execute(ReportRequestCurrentYear, LastSelectedTenantId)

    ' Initialize a new WebRequest object for the API request for the PreviousYear Trial Balance
    Dim ReportRequestPreviousYear As WebRequest
    Set ReportRequestPreviousYear = New WebRequest
    'ReportRequestPreviousYear.Resource = "api.xro/2.0/Reports/TrialBalance?Date=PreviousYear-06-30"
    ReportRequestPreviousYear.Resource = "api.xro/2.0/Reports/TrialBalance?Date=" & previousYear & "-06-30"
    ReportRequestPreviousYear.Method = WebMethod.HttpGet
    ReportRequestPreviousYear.RequestFormat = WebFormat.FormUrlEncoded
    ReportRequestPreviousYear.ResponseFormat = WebFormat.Json

    ' Send the API request for PreviousYear data and retrieve the response
    Dim ReportResponsePreviousYear As WebResponse
    Set ReportResponsePreviousYear = XeroClient.Execute(ReportRequestPreviousYear, LastSelectedTenantId)
    
    ' Check if all requests were successful and combine them into a dictionary
    If AccountsResponse.StatusCode = WebStatusCode.Ok And ReportResponseCurrentYear.StatusCode = WebStatusCode.Ok And ReportResponsePreviousYear.StatusCode = WebStatusCode.Ok Then
        ' Add the Chart of Accounts and Trial Balance data to the combined dictionary
        Dim CombinedReport As New Dictionary
        CombinedReport.Add "Accounts", AccountsResponse.Data
        CombinedReport.Add "CurrentYear", ReportResponseCurrentYear.Data
        CombinedReport.Add "PreviousYear", ReportResponsePreviousYear.Data
        Set GetChartOfAccounts = CombinedReport
    Else
        ' Handle API request failure
        Err.Raise 11041 + vbObjectError, "XeroAPICall.GetCombinedReportData", "Failed to retrieve either chart of accounts or trial balance data."
    End If

ApiCall_Cleanup:
    ' Clean up objects and handle errors
    Set AccountsRequest = Nothing
    Set AccountsResponse = Nothing
    Set ReportRequestCurrentYear = Nothing
    Set ReportResponseCurrentYear = Nothing
    Set ReportRequestPreviousYear = Nothing
    Set ReportResponsePreviousYear = Nothing
End Function





''
' Parses response data (JSON in the form of a Dictionary) and loads it into an Excel sheet.
'
' @method LoadReportToSheet
' @param {Dictionary} GetReportData - The JSON object obtained from an API call.
' @param {String} [SheetName=ReportOutputSheet] - Optional name for the output sheet.
'
' This function performs the following steps:
' 1. Extracts the report and its components from the JSON response.
' 2. Initializes the Excel sheet and sets the starting row index.
' 3. Writes the report titles to the sheet, formatting them appropriately.
' 4. Adds the account and date headers, formatting the date header based on the report period.
' 5. Iterates over the sections and rows of the report, adding data to the sheet and applying styles.
' 6. Adjusts the sheet name to avoid duplicates and applies final formatting.
' 7. Handles any errors that occur and logs them.
''

Private Sub LoadReportToSheet(GetReportData As Dictionary, Optional SheetName As String = ReportOutputSheet)
    On Error GoTo ApiCall_Cleanup

    ' Extract the CurrentYear and PreviousYear report data from the JSON object
    Dim reportCurrentYear As Dictionary, reportPreviousYear As Dictionary
    Set reportCurrentYear = GetReportData("CurrentYear")("Reports")(1)
    Set reportPreviousYear = GetReportData("PreviousYear")("Reports")(1)
    
    Dim reportTitlesCurrentYear As Collection
    Set reportTitlesCurrentYear = reportCurrentYear("ReportTitles")
    
    Dim rowsCurrentYear As Collection, rowsPreviousYear As Collection
    Set rowsCurrentYear = reportCurrentYear("Rows")
    Set rowsPreviousYear = reportPreviousYear("Rows")
    
    ' Initialize the row index for the Excel sheet
    Dim rowIndex As Long
    rowIndex = 1 ' Start at the first row

    ' Create a new sheet to load the JSON data into
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    
    ' Write the report titles to the sheet (use the same titles for both reports)
    With sh
        Dim reportTitle As Variant
        
        ' Add titles for CurrentYear (starting from A1)
        For Each reportTitle In reportTitlesCurrentYear
            .Cells(rowIndex, 1).value = reportTitle
            rowIndex = rowIndex + 1
        Next reportTitle
        
        ' Add a blank row after the report titles
        rowIndex = rowIndex + 1
        
        ' Add column headers (Account Number, Account Name, Account Type, etc.)
        .Cells(rowIndex, 1).value = "Account Number"
        .Cells(rowIndex, 2).value = "Account Name"
        .Cells(rowIndex, 3).value = "Account Type"
        .Cells(rowIndex, 4).value = "Current Year: Debit"
        .Cells(rowIndex, 5).value = "Current Year: Credit"
        .Cells(rowIndex, 6).value = "Previous Year: Debit - Credit"
        .Cells(rowIndex, 7).value = "Account Type from TrialBalance"
        .Cells(rowIndex, 8).value = "Maching Account from ChartofAccount" ' New column for manually mapped Type
        .Cells(rowIndex, 9).value = "AccountID"
        ' Add a blank row after the headers
        rowIndex = rowIndex + 1
        
        ' Process CurrentYear and PreviousYear rows together
        Dim sectionCurrentYear As Dictionary, sectionPreviousYear As Dictionary
        Dim innerRowsCurrentYear As Collection, innerRowsPreviousYear As Collection
        Dim rowCurrentYear As Dictionary, rowPreviousYear As Dictionary
        
        Dim i As Integer, j As Integer
        
        ' Start processing the CurrentYear data
        For Each sectionCurrentYear In rowsCurrentYear
            If sectionCurrentYear("RowType") = "Section" Then
                If sectionCurrentYear("Title") <> vbNullString Then
                    ' Insert section title (e.g., "Revenue", "Expenses", etc.)
                    .Cells(rowIndex, 1).value = sectionCurrentYear("Title")
                    rowIndex = rowIndex + 1
                End If
                
                Set innerRowsCurrentYear = sectionCurrentYear("Rows")
                
                For i = 1 To innerRowsCurrentYear.Count
                    Set rowCurrentYear = innerRowsCurrentYear(i)
                    
                    ' Skip "Total" row
                    If rowCurrentYear("Cells")(1)("Value") = "Total" Then
                        Exit For
                    End If
                    
                    ' Extract Account Name and Account Number
                    Dim accountName As String
                    Dim accountNumber As String
                    Dim accountID As String
                    
                    accountName = rowCurrentYear("Cells")(1)("Value")
                    accountID = rowCurrentYear("Cells")(1)("Attributes")(1)("Value")
                    accountNumber = vbNullString ' Initialize as empty
                    
                    ' Find the position of the last set of parentheses
                    Dim lastOpenParen As Long
                    Dim lastCloseParen As Long
                    
                    lastOpenParen = InStrRev(accountName, "(")
                    lastCloseParen = InStrRev(accountName, ")")
                    
                    ' Check if Account Name contains a number in the last set of parentheses
                    If lastOpenParen > 0 And lastCloseParen > lastOpenParen Then
                        ' Extract the data from within the parentheses
                    accountNumber = Trim(Mid(accountName, lastOpenParen + 1, lastCloseParen - lastOpenParen - 1))
                    End If
                    
                    ' Write Account Type, Account Name, and Account Number to sheet
                    .Cells(rowIndex, 1).value = "'" & accountNumber ' Write the extracted Account Number
                    .Cells(rowIndex, 2).value = accountName ' Account Name
                    .Cells(rowIndex, 7).value = sectionCurrentYear("Title")  ' Account Type
                    .Cells(rowIndex, 4).value = rowCurrentYear("Cells")(4)("Value") ' CurrentYear Debit YTD
                    .Cells(rowIndex, 5).value = rowCurrentYear("Cells")(5)("Value") ' CurrentYear Credit YTD
                    .Cells(rowIndex, 9).value = accountID
                    
                    ' Process corresponding PreviousYear row
                    ' Find matching account in PreviousYear data by account name
                    For Each sectionPreviousYear In rowsPreviousYear
                        If sectionPreviousYear("RowType") = "Section" Then
                            Set innerRowsPreviousYear = sectionPreviousYear("Rows")
                            For j = 1 To innerRowsPreviousYear.Count
                                Set rowPreviousYear = innerRowsPreviousYear(j)
                                ' Compare account names to find the matching row in PreviousYear
                                If rowCurrentYear("Cells")(1)("Value") = rowPreviousYear("Cells")(1)("Value") Then
                                    ' Calculate Debit - Credit for PreviousYear and write result to column 6
                                    Dim debitValue As Double
                                    Dim creditValue As Double
                                    
                                    ' If Debit or Credit is empty, treat as 0
                                    debitValue = IIf(IsNumeric(rowPreviousYear("Cells")(4)("Value")), rowPreviousYear("Cells")(4)("Value"), 0)
                                    creditValue = IIf(IsNumeric(rowPreviousYear("Cells")(5)("Value")), rowPreviousYear("Cells")(5)("Value"), 0)
                                    
                                    ' Write the result of Debit - Credit into column 6
                                    .Cells(rowIndex, 6).value = debitValue - creditValue
                                    
                                    Exit For
                                End If
                            Next j
                        End If
                    Next sectionPreviousYear
                    
                    rowIndex = rowIndex + 1
                Next i
            End If
        Next sectionCurrentYear
        
        ' After processing all CurrentYear data, process any missing PreviousYear accounts
        Dim accountFound As Boolean
        
        ' Iterate through PreviousYear sections
        For Each sectionPreviousYear In rowsPreviousYear
            If sectionPreviousYear("RowType") = "Section" Then
                Set innerRowsPreviousYear = sectionPreviousYear("Rows")
                
                ' Check each account in PreviousYear
                For j = 1 To innerRowsPreviousYear.Count
                    Set rowPreviousYear = innerRowsPreviousYear(j)
                    
                    accountFound = False ' Assume the account is missing
                    
                    ' Compare this account with all CurrentYear accounts
                    For Each sectionCurrentYear In rowsCurrentYear
                        If sectionCurrentYear("RowType") = "Section" Then
                            Set innerRowsCurrentYear = sectionCurrentYear("Rows")
                            For i = 1 To innerRowsCurrentYear.Count
                                Set rowCurrentYear = innerRowsCurrentYear(i)
                                If rowCurrentYear("Cells")(1)("Value") = rowPreviousYear("Cells")(1)("Value") Then
                                    ' Account is found in CurrentYear, no need to add from PreviousYear
                                    accountFound = True
                                    Exit For
                                End If
                            Next i
                        End If
                        If accountFound Then Exit For
                    Next sectionCurrentYear
                    
                    ' If the account was not found in CurrentYear, add it to the sheet
                    If Not accountFound Then
                        ' Extract Account Name and Account Number for PreviousYear account
                        Dim accountNamePreviousYear As String
                        Dim accountNumberPreviousYear As String
                        Dim accountIDPreviousYear As String
                        
                        accountNamePreviousYear = rowPreviousYear("Cells")(1)("Value")
                        accountIDPreviousYear = rowPreviousYear("Cells")(1)("Attributes")(1)("Value")
                        accountNumberPreviousYear = vbNullString ' Initialize as empty
                        
                        ' Find the position of the last set of parentheses
                        lastOpenParen = InStrRev(accountNamePreviousYear, "(")
                        lastCloseParen = InStrRev(accountNamePreviousYear, ")")
                        
                        ' Check if Account Name contains a number in parentheses
                        If lastOpenParen > 0 And lastCloseParen > lastOpenParen Then
                            ' Extract the data from within the parentheses
                            accountNumberPreviousYear = Trim(Mid(accountNamePreviousYear, lastOpenParen + 1, lastCloseParen - lastOpenParen - 1))
                            ' Force the number to be treated as text by adding a single quote before it


                        End If
                        
                        ' Write the missing account information to the sheet
                        .Cells(rowIndex, 1).value = "'" & accountNumberPreviousYear ' Account Number
                        .Cells(rowIndex, 2).value = accountNamePreviousYear ' Account Name
                        .Cells(rowIndex, 7).value = sectionPreviousYear("Title") ' Account Type
                        .Cells(rowIndex, 9).value = accountIDPreviousYear ' Account ID
                        
                        ' Calculate and write the Debit - Credit for PreviousYear
                        Dim debitValuePreviousYear As Double
                        Dim creditValuePreviousYear As Double
                        
                        ' If Debit or Credit is empty, treat as 0
                        debitValuePreviousYear = IIf(IsNumeric(rowPreviousYear("Cells")(4)("Value")), rowPreviousYear("Cells")(4)("Value"), 0)
                        creditValuePreviousYear = IIf(IsNumeric(rowPreviousYear("Cells")(5)("Value")), rowPreviousYear("Cells")(5)("Value"), 0)
                        
                        .Cells(rowIndex, 6).value = debitValuePreviousYear - creditValuePreviousYear ' PreviousYear Debit - Credit
                        
                        ' Move to the next row for the next account
                        rowIndex = rowIndex + 1
                    End If
                Next j
            End If
        Next sectionPreviousYear
        
    End With

    ' Now we perform the pairing with ReportingCodeName after loading all the data
    ' Prepare a dictionary to map Account Code to ReportingCodeName
    Dim accountMapping As Object
    Set accountMapping = CreateObject("Scripting.Dictionary")
    
    ' Parse the Accounts data
    Dim accountsData As Collection
    Set accountsData = GetReportData("Accounts")("Accounts")
    Dim account As Dictionary
    For Each account In accountsData
        accountMapping(account("AccountID")) = account("Type")
    Next account
    
    ' Prepare the manually written mapping dictionary
    Dim manuallyMappedType As Object
    Set manuallyMappedType = CreateObject("Scripting.Dictionary")
    
    '' Add the mappings manually
    ' Asset
    manuallyMappedType.Add "BANK", "Bank"
    manuallyMappedType.Add "CURRENT", "Current Asset"
    manuallyMappedType.Add "FIXED", "Fixed Asset"
    manuallyMappedType.Add "INVENTORY", "Inventory"
    manuallyMappedType.Add "NONCURRENT", "Non-current Asset"
    manuallyMappedType.Add "PREPAYMENT", "Prepayment"

    
    ' Liability
    manuallyMappedType.Add "CURRLIAB", "Current Liability"
    manuallyMappedType.Add "TERMLIAB", "Non-current Liability"
    manuallyMappedType.Add "LIABILITY", "Liability"
    
    ' Equity
    manuallyMappedType.Add "EQUITY", "Equity"
    
    ' Revenue
    manuallyMappedType.Add "REVENUE", "Revenue"
    manuallyMappedType.Add "OTHERINCOME", "Other Income"
    manuallyMappedType.Add "SALES", "Sales"
    
    ' Expense
    manuallyMappedType.Add "DIRECTCOSTS", "Direct Costs"
    manuallyMappedType.Add "DEPRECIATN", "Depreciation"
    manuallyMappedType.Add "OVERHEAD", "Overhead"
    manuallyMappedType.Add "EXPENSE", "Expense"
    
    
    
    ' Iterate over the rows where AccountID is written, and match both the ReportingCodeName and manually mapped type
    Dim lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 9).End(xlUp).row ' Find the last row with data in column I (AccountID)
    
    Dim currentAccountNumber As String
    Dim matchingAccountType As String
    Dim k As Long
    For k = 7 To lastRow ' Assuming data starts from row 7 after headers
        currentAccountNumber = sh.Cells(k, 9).value
        
        ' Lookup the ReportingCodeName from the accountMapping dictionary
        If accountMapping.Exists(currentAccountNumber) Then
            matchingAccountType = accountMapping(currentAccountNumber)
            sh.Cells(k, 8).value = matchingAccountType
        Else
            matchingAccountType = "Not Found"
            sh.Cells(k, 8).value = matchingAccountType
        End If
        
        ' Lookup the manually mapped type from the manuallyMappedType dictionary
        If manuallyMappedType.Exists(matchingAccountType) Then
            sh.Cells(k, 3).value = manuallyMappedType(matchingAccountType)
        Else
            sh.Cells(k, 3).value = "Not Found"
        End If
    Next k


    ' Set the sheet name to avoid duplicates
    Dim sheetIndex As Long
    sheetIndex = 0
    If WebHelpers.WorksheetExists(SheetName, ThisWorkbook) Then
        Do While WebHelpers.WorksheetExists(SheetName & "_" & sheetIndex, ThisWorkbook)
            sheetIndex = sheetIndex + 1
        Loop
        SheetName = SheetName & "_" & sheetIndex
    End If
    sh.name = SheetName
    
    ' Turn off gridlines for better presentation
    WebHelpers.TurnOffGridLines sh
    
    ' Save the name of the sheet
    pLastOutputSheetName = SheetName

    ' Sort data from A5 based on the "Account Number" column in ascending order
    With sh.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sh.Range("A5"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sh.Range("A5").CurrentRegion
        .Header = xlYes
        .Apply
    End With
    
    
    ' Now copy the data to the "Trial Balance Dump" sheet
    Dim dumpSheet As Worksheet
    Set dumpSheet = ThisWorkbook.Sheets("Trial Balance Dump")
    
    ' Clear the contents of the "Trial Balance Dump" sheet before pasting
    dumpSheet.Cells.Clear
    
    ' Copy data from the temporary sheet to "Trial Balance Dump"
    sh.UsedRange.Copy
    dumpSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    ' Bold the first three rows
    dumpSheet.Range("A1:Z4").Font.Bold = True

    ' Change format to "#,##0.00;(#,##0.00)" from columns D to F
    dumpSheet.Range("D:F").NumberFormat = "#,##0.00;(#,##0.00)"
    
    ' ' Set the mapping account color to light blue
    ' dumpSheet.Columns(3).Interior.Color = RGB(173, 216, 230)
    ' dumpSheet.Columns(7).Interior.Color = RGB(173, 216, 230)
    ' dumpSheet.Columns(8).Interior.Color = RGB(173, 216, 230)
    
    ' Hide column 7 and 8
    dumpSheet.columns(7).Hidden = True
    dumpSheet.columns(8).Hidden = True
    dumpSheet.columns(9).Hidden = True
    
    ' Delete the temporary sheet after copying the data
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True

    ' Stay on the "Trial Balance Dump" Sheet
    dumpSheet.activate
    
ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while loading report to sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", vbNullString) & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.LoadReportToSheet", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAPICall.LoadReportToSheet", auth_ErrorDescription
    End If
End Sub



' --------------------------------------------- '
' Execution
' --------------------------------------------- '

''
' Calls the login procedures for the user interface button.
'
' @method Login_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Retrieves the pre-set authenticator object from the XeroClient.
' 3. Logs out and clears the cache for the current session.
' 4. Initiates the login process.
' 5. Returns the authenticator reference to the XeroClient.
' 6. Handles any errors that occur during the process and logs them.
''
Public Sub Login_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve the pre-set authenticator object
    Dim Auth As XeroAuthenticator
    Set Auth = XeroClient.Authenticator
    Set XeroClient.Authenticator = Nothing
    
    ' Logout and clear cache for the current session
    Auth.Logout
    
    ' Login
    Auth.Login
    
    ' Return the authenticator reference to the XeroClient
    Set XeroClient.Authenticator = Auth
    ' Clear the local reference to the authenticator
    Set Auth = Nothing
    
ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error happened
        pXeroClientId = vbNullString
        pXeroClientSecret = vbNullString
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the login process." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.Login_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub

''
' Calls the report generation procedures for the user interface button.
'
' @method GenerateReport_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Retrieves the Trial balance report from the API.  # Revise
' 3. Parses and loads the report data into an Excel sheet.
' 4. Displays a message box to notify the user of the successful report generation.
' 5. Handles any errors that occur during the process and logs them.
''
Public Sub GenerateReport_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve the chart of accounts from the API
    Dim ReportDict As Dictionary
    Set ReportDict = GetChartOfAccounts
    
    
    ' Parse and load the report data into an Excel sheet
    LoadReportToSheet ReportDict
    ' Notify the user of successful report generation
    MsgBox "Report successfully generated on Trial Balance Dump sheet"

ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error occurred
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while generating report." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.GenerateReport_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub


''
' Clears all saved tokens and Xero organizations/tenants ID for the user interface button.
'
' @method ClearCache_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Confirms the user's action to clear the cache.
' 3. If the user confirms, retrieves the pre-set authenticator object.
' 4. Clears all cache (tenants and tokens) and logs out of the current session.
' 5. Returns the authenticator reference to the XeroClient.
' 6. Handles any errors that occur during the process and logs them.
''
Public Sub ClearCache_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Confirm user action
    Dim msgBoxResponse As VbMsgBoxResult
    msgBoxResponse = MsgBox("This action will clear saved tokens (access) and Xero organization IDs. You will be required to log in for the next request to generate a report." & _
        vbNewLine & vbNewLine & "Proceed to clears cache?", vbExclamation + vbYesNo, "Xero Report Generator - Microsoft Excel")
    
    Select Case msgBoxResponse
        Case vbYes
            ' Retrieve the pre-set authenticator object
            Dim Auth As XeroAuthenticator
            Set Auth = XeroClient.Authenticator
            ' Clear the reference to the authenticator in the XeroClient
            Set XeroClient.Authenticator = Nothing
            
            ' Clear all cache (tenants and tokens)
            Auth.ClearAllCache isClearTenant:=True, isClearToken:=True
            
            ' Clear current session tokens cache by logging out
            Auth.Logout
            
            ' Return the authenticator reference to the XeroClient
            Set XeroClient.Authenticator = Auth
            ' Clear the local reference to the authenticator
            Set Auth = Nothing
            
        Case vbNo
            ' Exit the subroutine if the user cancels the action
            Exit Sub
    End Select

ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error occurred
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.ClearCache_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub


