# Workpaper Guidance



# 1. Save Files
### 1.1 Save Customer File Under G Drive
- When exporting files from Content Snare:
  - Avoid using the original attachment filename.
  - This ensures better matching of accounts.

### 1.2 Save Workpapers in a Separate Folder
- Organize workpapers under a dedicated folder for streamlined access.

---

### 2. Generate Xero Report (Using Excel - Xero Report Generator)

#### 2.1 Log In
1. Paste the Client ID into the appropriate field.
2. Provide the credentials to log in to your Xero account.
3. Connect to the organization that requires the report.

#### 2.2 Generate Report and Update
1. Click **Generate Report and Update**.
2. Select the desired organization.
3. The report will be generated and saved under the **Trial Balance Dump**.

---

### 3. Match Chart of Accounts with File Path (Using Excel - File Path)

#### 3.1 Set File Path
- Paste the folder path storing the document into cell **B2**.

#### 3.2 Check File Existence
1. Click the **Check if file exists** button.
   - Subfolder name, file name, and file path will appear in columns **G** to **I**.

#### 3.3 Select Related Sections
- Under column **D**, specify the section requiring source documents.

#### 3.4 Match Accounts
1. Click the **Match accounts** button.
   - Documents are matched first by account number, then by chosen sections.
2. Click **Check if file exists** again to find missing documents and manually match related accounts.

#### 3.5 Adjust Columns
- Use the **Hide/Unhide COL** button to manage visibility of columns **G**, **H**, and **J** for manual modifications.

#### 3.6 Close Excel
- Proceed to the Python script for further steps.

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

---

### 5. Edit in FYI
- Links to related documents will be auto-generated when selecting the Xero account.

---

## Notes
- Ensure proper folder organization and naming conventions for seamless integration with Xero and FYI.
- Regularly check for missing documents and update the matching process as needed.
