# Goal of this Project
This project leverages a VBA script to retrieve an entity's Xero trial balance through the API. 

Using a combination of VBA and Python scripts, the chart of accounts is matched with local files provided by the entity, and the documents are seamlessly uploaded to FYI.



 
# Blcks and Credits
This project was built in the VBA 7.1 programming language. It was made possible thanks to open-source modules/packages:

- [vba-xero-api ](https://github.com/Muyoouu/vba-xero-api?tab=readme-ov-file#readme-top) by Muyoouu. 
- [VBA-Web](https://github.com/VBA-tools/VBA-Web) by Tim Hall
- [Chromium Automation with CDP for VBA ](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA))by ChrisK23 & Long Vh




# Workpaper Guidance
## 1. Save Files
### 1.1 Save entity's File Under the Local File

### 1.2 Save Workpaper and Python Script in a Separate Folder
- Organize workpapers under a dedicated folder for streamlined access.


## 2. Generate Xero Report (Using Excel - Xero Report Generator)

### 2.1 Log In
1. Paste the Client ID into the appropriate field.
2. Provide the credentials to log in to your Xero account.
3. Connect to the organization that requires the report.

### 2.2 Generate Report and Update
1. Click **Generate Report and Update**.
2. Select the desired organization.
3. The report will be generated and saved under the **Trial Balance Dump**.
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Select Org.png" alt="Select Organization" title="Select Organization" width="700">

<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Choose Year.png" alt="Choose year of the Organization" title="Choose year of the Organization" width="700">

## 3. Match Chart of Accounts with File Path (Using Excel - File Path)

### 3.1 Set File Path
- Paste the folder path storing the document into cell **B2**.

### 3.2 Check File Existence
1. Click the **Check if file exists** button.
   - Subfolder name, file name, and file path will appear in columns **G** to **I**.

### 3.3 Select Related Sections
- Under column **D**, specify the section requiring source documents.

### 3.4 Match Accounts
1. Click the **Match accounts** button.
   - Documents are matched first by account number, then by chosen sections.
2. Click **Check if file exists** again to find missing documents and manually match related accounts.

### 3.5 Adjust Columns
- Use the **Hide/Unhide COL** button to manage visibility of columns **G**, **H**, and **J** for manual modifications.

### 3.6 Close Excel
- Proceed to the Python script for further steps.
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Macth account with local files.png" alt="Match account with local files" title="Match account with local files" width="700">
---

### 4. Upload Documents and Worksheet (Using Python)

#### 4.1 Run Python Script
1. Click **Run All** in the script.
   - The code will upload:
     - Local documents.
     - Excel files renamed as **2024 Annual Workpapers + Customer Name**.

#### 4.2 Handle Duplicates
- If duplicate customers exist, a selection window will appear for entity selection.

#### 4.3 Access Worksheet in FYI
- Click the generated link to navigate to the corresponding worksheet in FYI.

<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Upload in FYI.png" alt="Upload in FYI" title="Upload in FYI" width="700">

### 5. Edit in FYI
- Links to related documents will be auto-generated when selecting the Xero account.
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Preview Link.png" alt="Preview Link" title="Preview Link" width="700">

## Notes
- Ensure proper folder organization and naming conventions for seamless integration with Xero and FYI.
- Regularly check for missing documents and update the matching process as needed.
