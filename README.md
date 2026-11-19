# 📊 CRM Data Cleaning & Management

### Improving CRM Data Accuracy Through Microsoft Excel

A practical Sales Operations project focused on cleaning, validating, standardising and organising CRM records to improve data accuracy and support reliable sales reporting.

**Business Domain:** Sales Operations | CRM Management | Data Quality  
**Target Market:** United States  
**Dataset:** 30 fictional CRM records  
**Tool:** Microsoft Excel

---

## 📈 Project Dashboard

<table>
<tr>
<td align="center">
<h2>30</h2>
<b>CRM Records</b><br>
<sub>Records reviewed and processed</sub>
</td>
<td align="center">
<h2>0</h2>
<b>Missing Values</b><br>
<sub>No blank fields identified in the checks</sub>
</td>
<td align="center">
<h2>1</h2>
<b>Invalid Email</b><br>
<sub>Initially identified and corrected</sub>
</td>
<td align="center">
<h2>30</h2>
<b>Validated Records</b><br>
<sub>Marked Good after correction</sub>
</td>
</tr>
</table>

*Metrics reflect the completed checks on the project's fictional dataset.*

---

## 📋 1. Project Overview

| Attribute | Details |
|---|---|
| Project Name | CRM Data Cleaning & Management |
| Project Type | Sales Operations |
| Business Scenario | Fictional B2B SaaS Company |
| Target Market | United States |
| Dataset Size | 30 CRM Records |
| Primary Tool | Microsoft Excel |
| Project Focus | Data Cleaning, Validation and Standardisation |
| Final Deliverable | Cleaned CRM Database and Data Quality Report |

### Business Context

CRM systems contain important information about prospects, customers and business accounts. However, inconsistent names, invalid contact details, duplicate records and missing information can affect sales activities.

This project simulates a CRM data management process in which raw prospect information is reviewed, validated, corrected and prepared for business use.

---

## 🎯 2. Project Objectives

- Organise raw CRM records into a structured database.
- Identify potential duplicate email addresses.
- Check for missing information.
- Standardise first names, last names and company names.
- Validate email address formats.
- Validate US phone number formats.
- Identify records requiring further review.
- Prepare a cleaned CRM database.
- Create a Data Quality Report to document the checks and corrections.

---

## 🔄 3. CRM Data Cleaning Workflow

The project followed a structured data management process.

<table>
<tr>
<td align="center" width="20%">
<h3>01</h3>
<b>Raw Data</b><br>
<sub>Original CRM records</sub>
</td>
<td align="center">→</td>
<td align="center" width="20%">
<h3>02</h3>
<b>Data Review</b><br>
<sub>Identify data issues</sub>
</td>
<td align="center">→</td>
<td align="center" width="20%">
<h3>03</h3>
<b>Data Cleaning</b><br>
<sub>Correct inconsistencies</sub>
</td>
<td align="center">→</td>
<td align="center" width="20%">
<h3>04</h3>
<b>Validation</b><br>
<sub>Check data accuracy</sub>
</td>
<td align="center">→</td>
<td align="center" width="20%">
<h3>05</h3>
<b>Final Database</b><br>
<sub>Cleaned CRM records</sub>
</td>
</tr>
</table>

### Workflow Explanation

| Stage | Activities | Output |
|---|---|---|
| 1. Raw Data Collection | Created 30 fictional CRM records | Raw CRM Data |
| 2. Data Review | Checked duplicates and missing values | Data Quality Findings |
| 3. Data Cleaning | Standardised names and company names | Cleaned Fields |
| 4. Data Validation | Checked email and phone formats | Validation Results |
| 5. Final Review | Combined validation checks and reviewed records | Data Quality Status |
| 6. Final Output | Prepared cleaned records and summary report | Cleaned CRM Database |

---

## 🧹 4. Data Cleaning Methodology

The following checks were performed during the project.

| Data Quality Check | Method Used | Purpose |
|---|---|---|
| Duplicate Check | COUNTIF | Identify duplicate email addresses |
| Missing Values | Go To Special → Blanks | Identify empty cells |
| Company Name Cleaning | TRIM | Remove unnecessary spaces |
| First Name Standardisation | PROPER | Standardise capitalisation |
| Last Name Standardisation | PROPER | Standardise capitalisation |
| Email Validation | IF, AND, ISNUMBER, SEARCH | Check basic email format |
| Phone Validation | SUBSTITUTE, LEN | Check US phone number format |
| Overall Data Quality | IF, AND | Assign record quality status |

### Examples of Data Cleaning

| Data Field | Before | After |
|---|---|---|
| First Name | JOHN | John |
| Last Name | SMITH | Smith |
| Company Name | BrightPath Digital (with extra spaces) | BrightPath Digital |
| Email | Invalid format | Corrected format |
| Phone | +1 (202) 555-0101 | Validated 11-digit format |

*The examples illustrate the types of corrections performed, not additional records.*

---

## 🧮 5. Excel Functions Used

| Excel Function | Application | Business Purpose |
|---|---|---|
| COUNTIF | Duplicate checking | Identify repeated email addresses |
| TRIM | Company name cleaning | Maintain consistent company names |
| PROPER | Name standardisation | Improve data consistency |
| IF | Conditional validation | Assign validation results |
| AND | Multiple condition checks | Evaluate several conditions together |
| ISNUMBER | Email validation | Check for required characters |
| SEARCH | Email validation | Locate characters within email addresses |
| SUBSTITUTE | Phone cleaning | Remove formatting characters |
| LEN | Phone validation | Check the number of digits |

### Example Formulas

**1. Company Name Cleaning**
```excel
=TRIM(G2)
```

**2. First Name Standardisation**
```excel
=PROPER(B2)
```

**3. Last Name Standardisation**
```excel
=PROPER(C2)
```

**4. Email Format Validation**
```excel
=IF(AND(ISNUMBER(SEARCH("@",G2)),ISNUMBER(SEARCH(".",G2))),"Valid","Invalid")
```

**5. Overall Data Quality Status**
```excel
=IF(AND(T2="Valid",U2="Valid",P2=1),"Good","Needs Review")
```

*Note: The formulas use the column references from the working dataset. Email validation checks basic formatting; it does not confirm whether an email address exists or can receive messages.*

---

## 📊 6. Data Quality Results

| Quality Check | Initial Finding | Final Result |
|---|---|---|
| Total Records | 30 | 30 |
| Missing Values | None identified | No missing values found |
| Duplicate Email Check | Checked | No duplicates flagged |
| Email Validation | 1 invalid email | Corrected |
| Phone Validation | 30 records checked | 30 valid |
| Overall Data Quality | 1 record required review | 30 records marked Good |

### Key Findings

- All 30 CRM records were reviewed.
- No missing values were identified in the checked fields.
- One invalid email format was identified and corrected.
- Phone numbers were validated using an 11-digit US number check.
- The initial review identified one record requiring attention.
- After correction, all 30 records were marked as Good.

---

## 📁 7. Workbook Structure

The Excel workbook is organised into four sheets.

| Sheet Name | Description |
|---|---|
| Raw CRM Data | Original CRM dataset retained for reference |
| Working CRM Data | Formulas, cleaning activities and validation checks |
| Cleaned CRM Data | Final corrected CRM records |
| Data Quality Report | Summary of data quality issues and actions taken |

### Data Fields

The CRM database contains the following information:

| Category | Fields |
|---|---|
| Record Information | Record ID |
| Contact Information | First Name, Last Name, Job Title, Email, Phone |
| Company Information | Company Name, Industry, City, Country, Company Size |
| Sales Information | Lead Source, Lead Status, Assigned Owner |
| Record Management | Date Added |

---

## 🛠️ 8. Tools Used

| Tool | Purpose |
|---|---|
| Microsoft Excel | CRM data organisation and management |
| Excel Formulas | Data cleaning and validation |
| Excel Tables | Structured data presentation and filtering |

---

## 📦 9. Project Deliverables

| Deliverable | Description | Status |
|---|---|---|
| Raw CRM Dataset | 30 fictional CRM records | Completed |
| Working CRM Dataset | Data cleaning and validation checks | Completed |
| Cleaned CRM Database | Corrected and standardised records | Completed |
| Data Quality Report | Summary of findings and corrections | Completed |
| Excel Workbook | Consolidated CRM management project | Completed |

---

## 💼 10. Skills Demonstrated

| Professional Category | Skills |
|---|---|
| CRM Management | CRM Data Organisation, Record Management |
| Sales Operations | Data Quality Checks, Sales Data Management |
| Data Cleaning | Data Standardisation, Duplicate Checking |
| Data Validation | Email Validation, Phone Validation |
| Microsoft Excel | COUNTIF, IF, AND, TRIM, PROPER, SEARCH, SUBSTITUTE, LEN |
| Reporting | Data Quality Reporting, Documentation |
| Process Management | Data Review, Error Identification, Data Correction |

---

## 💡 11. Project Outcomes & Key Learnings

Through this project, I gained practical experience in:

- Working with structured CRM datasets.
- Identifying common CRM data quality issues.
- Cleaning inconsistent company and contact information.
- Using Excel formulas to automate basic validation checks.
- Combining multiple validation conditions to identify records requiring review.
- Maintaining separate raw, working and final datasets.
- Documenting data quality findings in a structured report.

### Business Value

Maintaining accurate CRM records can help sales teams:

- Reduce duplicate prospect records.
- Improve contact information quality.
- Maintain consistent company and prospect details.
- Support more reliable CRM reporting.
- Improve visibility into lead management activities.

---

## 🚀 12. Future Improvements

| Priority | Improvement | Status |
|---|---|---|
| High | Add automated data quality checks for new records | Proposed |
| High | Introduce dropdown lists for standardised lead statuses | Proposed |
| Medium | Add conditional formatting for records requiring review | Proposed |
| Medium | Create a CRM data quality dashboard | Proposed |
| Low | Explore CRM data import and export workflows | Future Scope |

---

## 📝 13. Project Conclusion

This project demonstrates how Microsoft Excel can be used to clean, validate, standardise and organise CRM data through a structured data management process.

By working with a fictional B2B SaaS CRM dataset, I practised identifying data quality issues, applying Excel formulas, correcting inaccurate information and preparing a final cleaned database.

The project helped strengthen my practical understanding of CRM management, data quality and Sales Operations processes.
