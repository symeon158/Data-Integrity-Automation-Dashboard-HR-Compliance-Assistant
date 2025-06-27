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

![image](https://github.com/user-attachments/assets/e139a3d7-70ef-4667-9605-045dc5a8dee0)
![image](https://github.com/user-attachments/assets/fcea4581-721a-4285-ad73-755df05d565c)
![image](https://github.com/user-attachments/assets/393cae79-2952-424c-be8a-10dd622b7ad3)




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

### 📜 Office Script Code

```ts
function main(workbook: ExcelScript.Workbook): string {
  type Cell = string | number | boolean | null;
  interface ManagerGroup {
    ManagerEmail: string;
    Records: Record<string, Cell>[];
  }

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
      if (
        dataVals.some(row =>
          row[ci] != null && String(row[ci]).toUpperCase().includes(key)
        )
      ) {
        return renameMap[key];
      }
    }
    return orig;
  });

  const colIndex = (name: string) => renamedHeaders.indexOf(name);
  const companyColIndex = colIndex("Company");

  const companySheet = workbook.getWorksheet("Companies");
  const companyRange = companySheet.getRange("A2:A" + companySheet.getUsedRange().getRowCount());
  const companyValues = companyRange.getValues().flat().map(v => String(v).trim()).filter(v => v !== "");
  const allowedCompanies = new Set(companyValues);

  const filtered = dataVals.filter(r => {
    const comp = String(r[companyColIndex] || "").trim();
    return allowedCompanies.has(comp);
  });

  const mgrTable = workbook.getTable("Table3");
  const mgrHdrs = (mgrTable.getHeaderRowRange().getValues() as string[][])[0].map(h => h.trim());
  const mgrBody = mgrTable.getRangeBetweenHeaderAndTotal().getValues() as string[][];

  const compI = mgrHdrs.indexOf("Company");
  const emailI = mgrHdrs.indexOf("ManagerEmail");
  if (compI < 0 || emailI < 0) throw new Error("Table3 must have 'Company' and 'ManagerEmail'");

  const mgrMap = new Map<string, string>();
  for (const row of mgrBody) {
    const comp = String(row[compI] || "").trim();
    const mail = String(row[emailI] || "").trim();
    if (comp && mail) mgrMap.set(comp, mail);
  }

  const outHeaders = [...renamedHeaders, "ManagerEmail"];
  const outRows = filtered.map(r => {
    const comp = String(r[companyColIndex] || "").trim();
    const mail = mgrMap.get(comp) || "";
    return [...r, mail] as Cell[];
  });

  const nameDobSeen = new Map<string, number>();
  const managerEmailCounts = new Map<string, Map<string, number>>();

  const mgrColIdx = outHeaders.indexOf("ManagerEmail");
  const emailColIdx = colIndex("email");

  outRows.forEach(row => {
    const surname = String(row[colIndex("Surname")] || "").trim();
    const name = String(row[colIndex("Name")] || "").trim();
    const dob = String(row[colIndex("Date of Birth")] || "").trim();
    const key = `${surname}|${name}|${dob}`;
    nameDobSeen.set(key, (nameDobSeen.get(key) || 0) + 1);

    const personEmail = String(row[emailColIdx] || "").trim().toLowerCase();
    const manager = String(row[mgrColIdx] || "").trim();
    if (!managerEmailCounts.has(manager)) {
      managerEmailCounts.set(manager, new Map());
    }
    const mgrMapRef = managerEmailCounts.get(manager)!;
    if (personEmail !== "" && personEmail !== "null") {
      mgrMapRef.set(personEmail, (mgrMapRef.get(personEmail) || 0) + 1);
    }
  });

  const grouping = new Map<string, Record<string, Cell>[]>();
  const dateCols = ["Hire Date", "Retire Date", "Date of Birth"];
  const dateColIndices = dateCols.map(col => outHeaders.indexOf(col));

  const importantCols = ["Date of Birth", "Hire Date", "Division", "Department", "City"];

  outRows.forEach(row => {
    const managerEmail = String(row[mgrColIdx] || "").trim() || "NoEmail";
    const rec: Record<string, Cell> = {};

    const surname = String(row[colIndex("Surname")] || "").trim();
    const name = String(row[colIndex("Name")] || "").trim();
    const dob = String(row[colIndex("Date of Birth")] || "").trim();
    const key = `${surname}|${name}|${dob}`;
    const isDuplicate = (nameDobSeen.get(key) || 0) > 1;

    const retireDate = row[colIndex("Retire Date")];
    const retireCause = row[colIndex("Retire Cause")];

    const isBlank = (v: Cell) => v === null || (typeof v === "string" && v.trim() === "");

    const cond1 = !isBlank(retireDate) && isBlank(retireCause);
    const cond2 = isBlank(retireDate) && importantCols.some(c => isBlank(row[colIndex(c)]));
    const cond3 = isDuplicate;

    const personEmail = String(row[emailColIdx] || "").trim().toLowerCase();
    const jobProperty = String(row[colIndex("Job Property")] || "").trim().toUpperCase();
    const isEmailValid = personEmail !== "" && personEmail !== "null";
    const mgrMapRef = managerEmailCounts.get(managerEmail);
    const cond4 = isEmailValid && mgrMapRef && mgrMapRef.get(personEmail)! > 1;
    const cond5 = jobProperty === "ADMINISTRATIVE" && !isEmailValid;

    const shouldInclude = cond1 || cond2 || cond3 || cond4 || cond5;

    rec["Cond1"] = cond1;
    rec["Cond2"] = cond2;
    rec["Cond3"] = cond3;
    rec["Cond4"] = cond4;
    rec["Cond5"] = cond5;
    rec["ShouldInclude"] = shouldInclude;

    let reason = "";
    if (cond1) reason += "Missing Retire Cause; ";
    if (cond2) reason += "Missing Important Field; ";
    if (cond3) reason += "Duplicate Name/DOB; ";
    if (cond4) reason += "Duplicate Email per Manager; ";
    if (cond5) reason += "Missing Email for ADMINISTRATIVE; ";
    rec["Reason"] = reason.trim();

    outHeaders.forEach((h, i) => {
      let val = row[i];
      if (dateColIndices.includes(i) && typeof val === "number") {
        const jsDate = new Date(Date.UTC(1899, 11, 30) + val * 86400000);
        rec[h] = jsDate.toISOString().slice(0, 10);
      } else {
        rec[h] = val;
      }
    });

    rec["IsDuplicate"] = cond3;

    const arr = grouping.get(managerEmail) || [];
    arr.push(rec);
    grouping.set(managerEmail, arr);
  });

  const groups: ManagerGroup[] = [];
  grouping.forEach((recs, email) => {
    groups.push({ ManagerEmail: email, Records: recs });
  });

  return JSON.stringify(groups);
}
