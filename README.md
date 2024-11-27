# Goal of this Project
This project leverages a VBA script to retrieve an entity's Xero trial balance through the API.
<br>

Using a combination of VBA and Python scripts, the chart of accounts is matched with local files provided by the entity, and the documents are seamlessly uploaded to FYI.

<br>

---

# Blocks and Credits
This project was built in the VBA 7.1 programming language. It was made possible thanks to open-source modules/packages:

- [vba-xero-api ](https://github.com/Muyoouu/vba-xero-api?tab=readme-ov-file#readme-top) by Muyoouu. 
- [VBA-Web](https://github.com/VBA-tools/VBA-Web) by Tim Hall.
- [Chromium Automation with CDP for VBA ](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA))by ChrisK23 & Long Vh.
<br>

---
# Workpaper Guidance
## 1. Download the Files and Save in a Seperate Folder
- Python: [Upload Document](https://github.com/JacintaQ/Xero-Report-and-FYI-Upload/blob/main/UploadDocument.ipynb)  
- Excel: [Year Annual Workpapers](https://github.com/JacintaQ/Xero-Report-and-FYI-Upload/blob/main/YEAR%20Annual%20Workpapers.xlsm)
<br>


## 2. Generate Xero Report (Excel)

### 2.1 Log In
1. Click **Log in**.
2. Paste the Client ID and the credentials into the appropriate fields.
3. Connect to the organization that requires the report.
   - Steps creating the Xero API could be found: [Xero - Create API](https://github.com/JacintaQ/Xero-Report-and-FYI-Upload/blob/main/Xero%20-%20Create%20API.docx).

<img src="https://github.com/JacintaQ/Xero-Report-and-FYI-Upload/blob/main/img/Enter%20Client%20ID.png" alt="Select Organization" title="Select Organization" width="1000">  
<img src="https://github.com/JacintaQ/Xero-Report-and-FYI-Upload/blob/main/img/Client%20Secret.png" alt="Select Organization" title="Select Organization" width="1000">  
<br>

[(back to top)](#goal-of-this-project)

### 2.2 Generate Trial Balance Report
1. Click **Generate Report**.
2. Select the ending year and desired organization.
3. The report will be generated and saved under the **Trial Balance Dump**.
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Select Org.png" alt="Select Organization" title="Select Organization" width="1000">

<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Choose Year.png" alt="Choose year of the Organization" title="Choose year of the Organization" width="1000">
<br>

[(back to top)](#goal-of-this-project)

## 3. Match Chart of Accounts with File Path (Excel)

### 3.1 Set File Path
- Paste the folder path storing the document into cell **B2**.
  - Recommend seperate the documents with the workpapers to avoid conflicts.

### 3.2 Check File Existence
- Click the **Check if file exists** button.
   - Subfolder name, file name, and file path will appear in columns **G** to **I**.

### 3.3 Select Related Sections
- Under column **D**, specify the section requiring source documents.

### 3.4 Match Accounts
1. Click **Check if file exists** again to find missing documents and manually match related accounts.
2. Click the **Match accounts** button.
   - Documents are matched first by account number, then by chosen sections.
     
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Macth%20account%20with%20local%20files.png" alt="Match account with local files" title="Match account with local files" width="1000">

<br>

[(back to top)](#goal-of-this-project)

### 4. Upload Documents and Worksheet (Python)

#### 4.1 Run Python Script
1. Click **Run All** in the script.
   - The code will upload:
     - Local documents.
     - Excel files renamed as **2024 Annual Workpapers + Customer Name**.

#### 4.2 Handle Duplicates
- If duplicate customers exist, a selection window will appear for entity selection.

#### 4.3 Access Worksheet in FYI
- Click the generated link to navigate to the corresponding worksheet in FYI.

<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Upload in FYI.png" alt="Upload in FYI" title="Upload in FYI" width="1000">

<br>

[(back to top)](#goal-of-this-project)

### 5. Edit in FYI
- Links to related documents will be auto-generated when selecting the Xero account.
<img src="https://raw.githubusercontent.com/JacintaQ/Xero-Report-and-FYI-Upload/main/img/Preview Link.png" alt="Preview Link" title="Preview Link" width="1000">

<br>

[(back to top)](#goal-of-this-project)

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

<br>

[(back to top)](#goal-of-this-project)

