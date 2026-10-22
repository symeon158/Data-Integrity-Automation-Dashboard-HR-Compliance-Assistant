# 🧩 Data Integrity & Automation Dashboard – HR Compliance Assistant

## 📌 Project Overview
This project leverages **Office Scripts**, **Power Automate**, and **Power Query** to ensure **data quality**, **consistency**, and **automation** across company and subsidiary employee records.

## 🎯 Purpose
The solution was designed to **assist HR managers and data owners** in monitoring and improving the quality of employee data. It identifies:

- ✅ Missing or incomplete critical fields  
- 🧠 Duplicate entries (based on key identifiers like name and date of birth, emails)  
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
  - 📨 Embedded **HTML table** of the data which have problematic records 
- Automatically stores files in OneDrive

![image](https://github.com/user-attachments/assets/596579bd-6d78-4b23-95b8-899aa9493b4b)
<img width="190" height="713" alt="image" src="https://github.com/user-attachments/assets/1ba0d91f-251c-42a6-8ec8-a57240e42112" />







### 🧼 Power Query (Preprocessing)
Before automation begins, **Power Query** handles:
- Column renaming and unification
- Removal of null/irrelevant entries
- Standardization of values (e.g., dates, salaries)
- Column reordering for logical structure

---

## ✅ Benefits
- Ensures **HR data quality and completeness**
- Categorizes each issue with a clear **Reason** for follow-up
- Automatically informs the responsible managers
- Reduces manual review and Excel processing time
- Fully **scalable and maintainable** for all company entities

---

## 🔧 Technologies Used
- `Office Scripts (Excel TypeScript API)` – Dynamic row evaluation with **reason tagging logic**
- `Power Automate` (Flows, Dynamic Content, HTML email formatting) – Generates and sends tables **including the Reason column**
- `Power Query` (ETL logic, data cleaning) – Prepares clean data before applying rules
- `Excel Online` (Automated reporting)
- `Outlook 365 API` (Email delivery engine)

---

## 📁 Example Output
- JSON groupings by manager
- Clean CSV/Excel attachments including the **Reason** column per row
- Structured HTML summaries in emails, listing employees and their **reason for being flagged**

<img width="253" height="321" alt="image" src="https://github.com/user-attachments/assets/04ada4ed-150b-44e8-a5a5-bd8e59f6e97d" />

<img width="1193" height="119" alt="image" src="https://github.com/user-attachments/assets/5b888013-7251-4b62-9a44-7da2d19c2c46" />





### 📜 Office Script Code

```ts
function main(workbook: ExcelScript.Workbook): string {
  type Cell = string | number | boolean | null;
  interface ManagerGroup {
    ManagerEmail: string;
    ManagerName: string;
    Records: Record<string, Cell>[];
  }

  const isBlank = (v: Cell) => v === null || (typeof v === "string" && v.trim() === "");

  const sheet = workbook.getWorksheet("Fields");
  const table = sheet.getTables()[0];
  const headerVals = table.getHeaderRowRange().getValues() as string[][];
  const dataVals = table.getRangeBetweenHeaderAndTotal().getValues() as Cell[][];
  const headers = headerVals[0].map(h => h.trim());

  const renameMap: Record<string, string> = {
    "FEMALE": "Gender",
    "ΕΥΚΑΡΠΙΑ": "City",
    "ΑΟΡΙΣΤΟΥ ΧΡΟΝΟΥ": "Employment Relation",
    "DIVISION": "Division",
    "WOOD EFFECT": "Department"
  };

  const renamedHeaders = headers.map((orig, ci) => {
    for (const key of Object.keys(renameMap)) {
      if (dataVals.some(row => row[ci] != null && String(row[ci]).toUpperCase().includes(key))) {
        return renameMap[key];
      }
    }
    return orig;
  });

  const colIndex = (name: string) => renamedHeaders.indexOf(name);
  const companyColIndex = colIndex("Company");

  // Filter by allowed companies (sheet "Companies", col A)
  const companySheet = workbook.getWorksheet("Companies");
  const companyRange = companySheet.getRange("A2:A" + companySheet.getUsedRange().getRowCount());
  const companyValues = companyRange.getValues().flat().map(v => String(v).trim()).filter(v => v !== "");
  const allowedCompanies = new Set(companyValues);

  const filtered = dataVals.filter(r => {
    const comp = String(r[companyColIndex] || "").trim();
    return allowedCompanies.has(comp);
  });

  // Manager table "Table3": Company | ManagerEmail | Name
  const mgrTable = workbook.getTable("Table3");
  const mgrHdrs = (mgrTable.getHeaderRowRange().getValues() as string[][])[0].map(h => h.trim());
  const mgrBody = mgrTable.getRangeBetweenHeaderAndTotal().getValues() as string[][];
  const compI = mgrHdrs.indexOf("Company");
  const emailI = mgrHdrs.indexOf("ManagerEmail");
  const nameI = mgrHdrs.indexOf("Name");
  if (compI < 0 || emailI < 0 || nameI < 0) {
    throw new Error("Table3 must have 'Company', 'ManagerEmail' and 'Name' columns");
  }

  const mgrMap = new Map<string, { email: string; name: string }>();
  for (const row of mgrBody) {
    const comp = String(row[compI] || "").trim();
    const mail = String(row[emailI] || "").trim();
    const managerName = String(row[nameI] || "").trim();
    if (comp && mail) mgrMap.set(comp, { email: mail, name: managerName });
  }

  const managerNameMap = new Map<string, string>();
  mgrMap.forEach(info => { if (info.email) managerNameMap.set(info.email, info.name); });

  const outHeaders = [...renamedHeaders, "ManagerEmail", "ManagerName"];
  const outRows = filtered.map(r => {
    const comp = String(r[companyColIndex] || "").trim();
    const info = mgrMap.get(comp) || { email: "", name: "" };
    return [...r, info.email, info.name] as Cell[];
  });

  const nameDobSeen = new Map<string, number>();
  const managerEmailCounts = new Map<string, Map<string, number>>();

  const mgrColIdx = outHeaders.indexOf("ManagerEmail");

  // Email column: support both "email" and "Email"
  let emailColIdx = colIndex("email");
  if (emailColIdx < 0) emailColIdx = colIndex("Email");

  // Key for Cond3: duplicates are checked within the same company only
  const nameDobKey = (row: Cell[]) => {
    const company = String(row[companyColIndex] || "").trim();
    const surname = String(row[colIndex("Surname")] || "").trim();
    const name = String(row[colIndex("Name")] || "").trim();
    const dob = String(row[colIndex("Date of Birth")] || "").trim();
    return `${company}|${surname}|${name}|${dob}`;
  };

  // Pre-count Name/DOB duplicates (per company) and email duplicates (per manager, actives only)
  outRows.forEach(row => {
    const key = nameDobKey(row);
    nameDobSeen.set(key, (nameDobSeen.get(key) || 0) + 1);

    const isActive = isBlank(row[colIndex("Retire Date")]);
    const personEmail = String(row[emailColIdx] || "").trim().toLowerCase();
    const manager = String(row[mgrColIdx] || "").trim() || "NoEmail"; // same key as grouping below
    if (!managerEmailCounts.has(manager)) {
      managerEmailCounts.set(manager, new Map());
    }
    const mgrMapRef = managerEmailCounts.get(manager)!;
    if (isActive && personEmail && personEmail !== "null") {
      mgrMapRef.set(personEmail, (mgrMapRef.get(personEmail) || 0) + 1);
    }
  });

  const grouping = new Map<string, Record<string, Cell>[]>();

  const importantCols = [
    "Company", "Employee Id", "Surname", "Name", "Gender", "Employment Relation",
    "Job Property", "Date of Birth", "Hire Date", "Division", "Department",
    "Job Description", "City", "Supervisor Id", "Nominal Salary",
    "Tax No.", "Base Salary"
  ];

  outRows.forEach(row => {
    const managerEmail = String(row[mgrColIdx] || "").trim() || "NoEmail";
    const rec: Record<string, Cell> = {};

    // Cond3: duplicate Surname|Name|Date of Birth within the same company
    const isDuplicate = (nameDobSeen.get(nameDobKey(row)) || 0) > 1;

    const retireDate = row[colIndex("Retire Date")];
    const retireCause = row[colIndex("Retire Cause")];
    const isActive = isBlank(retireDate);

    // --- Cond7: Active must be consistent with Retire Date
    // Active column already exists in data; normalize it to 0/1 if possible
    let activeValNum: number | null = null;
    {
      const activeIdx = colIndex("Active");
      const activeRaw = activeIdx >= 0 ? row[activeIdx] : null;
      if (typeof activeRaw === "number") {
        activeValNum = activeRaw;
      } else if (typeof activeRaw === "boolean") {
        activeValNum = activeRaw ? 1 : 0;
      } else {
        const s = String(activeRaw ?? "").trim().toLowerCase();
        if (s === "1" || s === "0") activeValNum = Number(s);
        else if (s === "true") activeValNum = 1;
        else if (s === "false") activeValNum = 0;
        else activeValNum = null; // unknown
      }
    }
    const cond7 =
      (isActive && activeValNum !== 1) ||
      (!isActive && activeValNum !== 0);

    const missingFields = importantCols.filter(c => isBlank(row[colIndex(c)]));
    const cond1 =
      (!isBlank(retireDate) && isBlank(retireCause)) ||
      (isBlank(retireDate) && !isBlank(retireCause));
    const cond2 = isActive && missingFields.length > 0;
    const cond3 = isDuplicate;

    const personEmail = String(row[emailColIdx] || "").trim().toLowerCase();
    const jobProperty = String(row[colIndex("Job Property")] || "").trim().toUpperCase();
    const isEmailValid = personEmail !== "" && personEmail !== "null";
    const mgrMapRef = managerEmailCounts.get(managerEmail);

    // Cond4: duplicate email under the same manager — actives only
    const cond4 = isActive && isEmailValid && !!mgrMapRef && (mgrMapRef.get(personEmail) || 0) > 1;

    // Cond5: ADMINISTRATIVE without email — actives only
    const cond5 = isActive && jobProperty === "ADMINISTRATIVE" && !isEmailValid;

    // Cond6: non-"alumil" email requires Mobile Phone not blank — actives only
    const mobilePhone = String(row[colIndex("Mobile Phone")] || "").trim();
    const nonAlumil = !personEmail.includes("alumil"); // blank counts as non-alumil
    const cond6 = isActive && nonAlumil && mobilePhone === "";

    const shouldInclude = cond1 || cond2 || cond3 || cond4 || cond5 || cond6 || cond7;

    rec["Cond1"] = cond1;
    rec["Cond2"] = cond2;
    rec["Cond3"] = cond3;
    rec["Cond4"] = cond4;
    rec["Cond5"] = cond5;
    rec["Cond6"] = cond6;
    rec["Cond7"] = cond7; // Active vs Retire Date inconsistency
    rec["ShouldInclude"] = shouldInclude;
    rec["MissingFields"] = cond2 ? missingFields.join(", ") : "";

    const reasons: string[] = [];
    if (cond1) reasons.push("Inconsistent Retire Info (Date/Cause mismatch)");
    if (cond2) reasons.push("Missing Important Field");
    if (cond3) reasons.push("Duplicate Name/DOB in same Company");
    if (cond4) reasons.push("Duplicate Email");
    if (cond5) reasons.push("Missing Email for ADMINISTRATIVE");
    if (cond6) reasons.push("Missing Mobile Phone for non-Alumil email");
    if (cond7) reasons.push("Active flag inconsistent with Retire Date");
    rec["Reason"] = reasons.join("; ");

    // Copy data columns (format dates)
    const dateColIndices = ["Hire Date", "Retire Date", "Date of Birth"].map(col => outHeaders.indexOf(col));
    outHeaders.forEach((h, i) => {
      let val = row[i];
      if (dateColIndices.includes(i) && typeof val === "number") {
        const jsDate = new Date(Date.UTC(1899, 11, 30) + (val as number) * 86400000);
        rec[h] = jsDate.toISOString().slice(0, 10);
      } else {
        rec[h] = val;
      }
    });

    rec["ManagerName"] = managerNameMap.get(managerEmail) || "";
    rec["IsDuplicate"] = cond3;

    const arr = grouping.get(managerEmail) || [];
    arr.push(rec);
    grouping.set(managerEmail, arr);
  });

  const groups: ManagerGroup[] = [];
  grouping.forEach((recs, email) => {
    groups.push({
      ManagerEmail: email,
      ManagerName: managerNameMap.get(email) || "",
      Records: recs
    });
  });

  return JSON.stringify(groups);
}

