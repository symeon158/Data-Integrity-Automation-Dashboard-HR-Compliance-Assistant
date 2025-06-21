# 🧩 Data Integrity & Automation Dashboard – HR Compliance Assistant

## 📌 Project Overview
This project leverages **Office Scripts**, **Power Automate**, and **Power Query** to ensure **data quality**, **consistency**, and **automation** across company and subsidiary employee records.

## 🎯 Purpose
The solution was designed to **assist HR managers and data owners** in monitoring and improving the quality of employee data. It identifies:

- ✅ Missing or incomplete critical fields  
- 🧠 Duplicate entries (based on key identifiers like name and date of birth)  
- ⚠️ Inconsistent or incorrect retirement records  

It then notifies the relevant **company and subsidiary managers** by email with detailed reports.

---

## ⚙️ Key Components

### 🖋 Office Script
- Parses Excel data to:
  - Rename and normalize column headers
  - Filter data based on an approved company list
  - Append **manager email addresses**
  - Detect **duplicates** and **data gaps**
  - Group entries per manager and output structured JSON

### 🔄 Power Automate Flow
- Triggers from a manual button or schedule
- Processes JSON from the Office Script
- Filters entries using **custom rules**, such as:
  - Missing retirement cause with a retirement date
  - Incomplete key fields (e.g., Date of Birth, Hire Date)
  - Detected duplicate employee profiles
- Sends **customized email reports** per manager, including:
  - 📎 Attached `.csv` or `.xlsx` file  
  - 📨 Embedded **HTML table** of the data  
- Automatically stores files in OneDrive

### 🧼 Power Query (Preprocessing)
Before automation begins, **Power Query** handles:
- Column renaming and unification
- Removal of null/irrelevant entries
- Standardization of values (e.g., dates, salaries)
- Column reordering for logical structure

---

## ✅ Benefits
- Ensures **HR data quality and completeness**
- Automatically informs the responsible managers
- Reduces manual review and Excel processing time
- Fully **scalable and maintainable** for all company entities

---

## 🔧 Technologies Used
- `Office Scripts (Excel TypeScript API)`  
- `Power Automate` (Flows, Dynamic Content, HTML email formatting)  
- `Power Query` (ETL logic, data cleaning)  
- `Excel Online` (Automated reporting)  
- `Outlook 365 API` (Email delivery engine)

---

## 📁 Example Output
- JSON groupings by manager
- Clean CSV/Excel attachments
- Structured HTML summaries in emails

---

> ✨ _“A smart automation system for enforcing data discipline across the HR landscape.”_
