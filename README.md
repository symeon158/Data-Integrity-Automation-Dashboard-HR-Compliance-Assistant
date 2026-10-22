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

  // ---------- Helpers ----------
  const isBlank = (v: Cell | undefined): boolean =>
    v === null || v === undefined || (typeof v === "string" && v.trim() === "");

  const norm = (v: Cell | undefined): string => String(v ?? "").trim();

  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
  const DAY_MS = 86400000;

  // Converts an Excel date serial or a text date (yyyy-mm-dd or dd/mm/yyyy) to an Excel serial number
  const toSerial = (v: Cell | undefined): number | null => {
    if (typeof v === "number") return v;
    if (typeof v !== "string") return null;
    const s = v.trim();
    if (s === "") return null;
    let m: RegExpMatchArray | null = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return (Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) - EXCEL_EPOCH) / DAY_MS;
    m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
    if (m) return (Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])) - EXCEL_EPOCH) / DAY_MS;
    return null;
  };

  const todaySerial: number = Math.floor((Date.now() - EXCEL_EPOCH) / DAY_MS);
  const yearsBetween = (fromSerial: number, toSerialVal: number): number =>
    (toSerialVal - fromSerial) / 365.25;

  // Tax No.: keep digits only, restore leading zeros lost when stored as a number
  const normTaxNo = (v: Cell | undefined): string => {
    const digits = norm(v).replace(/\D/g, "");
    if (digits === "") return "";
    return digits.length < 9 ? ("000000000" + digits).slice(-9) : digits;
  };

  // Salary: number, or text in 1234.56 / 1234,56 / 1.234,56 formats; NaN if not a number
  const toNumber = (v: Cell | undefined): number => {
    if (typeof v === "number") return v;
    let s = norm(v).replace(/\s/g, "");
    if (s.indexOf(",") >= 0 && s.indexOf(".") >= 0) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(",", ".");
    return s === "" ? NaN : Number(s);
  };

  // ---------- Read main table ----------
  const sheet = workbook.getWorksheet("Fields");
  const table = sheet.getTables()[0];
  const headerVals = table.getHeaderRowRange().getValues() as string[][];
  const dataVals = table.getRangeBetweenHeaderAndTotal().getValues() as Cell[][];
  const headers: string[] = headerVals[0].map((h: string) => String(h).trim());

  const renameMap: Record<string, string> = {
    "FEMALE": "Gender",
    "ΕΥΚΑΡΠΙΑ": "City",
    "ΑΟΡΙΣΤΟΥ ΧΡΟΝΟΥ": "Employment Relation",
    "DIVISION": "Division",
    "WOOD EFFECT": "Department"
  };

  const renamedHeaders: string[] = headers.map((orig: string, ci: number) => {
    for (const key of Object.keys(renameMap)) {
      if (dataVals.some(row => row[ci] != null && String(row[ci]).toUpperCase().includes(key))) {
        return renameMap[key];
      }
    }
    return orig;
  });

  const colIndex = (name: string): number => renamedHeaders.indexOf(name);

  // ---------- Fail loudly if a required column is missing ----------
  const importantCols = [
    "Company", "Employee Id", "Surname", "Name", "Gender", "Employment Relation",
    "Job Property", "Date of Birth", "Hire Date", "Division", "Department",
    "Job Description", "City", "Supervisor Id", "Nominal Salary",
    "Tax No.", "Base Salary"
  ];
  const otherRequiredCols = ["Retire Date", "Retire Cause", "Active", "Mobile Phone"];
  const missingCols: string[] = [...importantCols, ...otherRequiredCols].filter(c => colIndex(c) < 0);

  let emailColIdx = colIndex("email");
  if (emailColIdx < 0) emailColIdx = colIndex("Email");
  if (emailColIdx < 0) missingCols.push("Email");

  if (missingCols.length > 0) {
    throw new Error("Missing column(s) in the 'Fields' table: " + missingCols.join(", "));
  }

  const companyColIndex = colIndex("Company");
  const empIdIdx = colIndex("Employee Id");
  const supIdIdx = colIndex("Supervisor Id");
  const retireIdx = colIndex("Retire Date");
  const taxIdx = colIndex("Tax No.");

  // ---------- Allowed companies (sheet "Companies", col A) ----------
  const companySheet = workbook.getWorksheet("Companies");
  const companyRange = companySheet.getRange("A2:A" + companySheet.getUsedRange().getRowCount());
  const companyValues: string[] = companyRange
    .getValues()
    .map((r: (string | number | boolean)[]) => String(r[0]).trim())
    .filter((v: string) => v !== "");
  const allowedCompanies = new Set<string>(companyValues);

  const filtered: Cell[][] = dataVals.filter(r => allowedCompanies.has(norm(r[companyColIndex])));

  // ---------- Managers (Table3: Company | ManagerEmail | Name) ----------
  const mgrTable = workbook.getTable("Table3");
  const mgrHdrs: string[] = (mgrTable.getHeaderRowRange().getValues() as string[][])[0].map((h: string) => String(h).trim());
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

  const outHeaders: string[] = [...renamedHeaders, "ManagerEmail", "ManagerName"];
  const outRows: Cell[][] = filtered.map(r => {
    const info = mgrMap.get(norm(r[companyColIndex])) || { email: "", name: "" };
    return [...r, info.email, info.name] as Cell[];
  });

  const mgrColIdx = outHeaders.indexOf("ManagerEmail");

  // ---------- Employee lookup for supervisor checks (all companies, incl. not-allowed ones) ----------
  // value = true if at least one record with that Id is active
  const empActiveByCompanyId = new Map<string, boolean>();
  const empActiveById = new Map<string, boolean>();
  for (const r of dataVals) {
    const id = norm(r[empIdIdx]);
    if (id === "") continue;
    const active = isBlank(r[retireIdx]);
    const k = norm(r[companyColIndex]) + "|" + id;
    empActiveByCompanyId.set(k, (empActiveByCompanyId.get(k) || false) || active);
    empActiveById.set(id, (empActiveById.get(id) || false) || active);
  }

  // ---------- Pre-counts ----------
  const nameDobSeen = new Map<string, number>();
  const taxSeen = new Map<string, number>();
  const empIdSeen = new Map<string, number>();
  const managerEmailCounts = new Map<string, Map<string, number>>();

  // Cond3 key: duplicates within the same company only
  const nameDobKey = (row: Cell[]): string => {
    const company = norm(row[companyColIndex]);
    const surname = norm(row[colIndex("Surname")]);
    const name = norm(row[colIndex("Name")]);
    const dob = norm(row[colIndex("Date of Birth")]);
    return `${company}|${surname}|${name}|${dob}`;
  };
  const taxKey = (row: Cell[]): string => {
    const t = normTaxNo(row[taxIdx]);
    return t === "" ? "" : norm(row[companyColIndex]) + "|" + t;
  };
  const empIdKey = (row: Cell[]): string => {
    const id = norm(row[empIdIdx]);
    return id === "" ? "" : norm(row[companyColIndex]) + "|" + id;
  };

  outRows.forEach(row => {
    const key = nameDobKey(row);
    nameDobSeen.set(key, (nameDobSeen.get(key) || 0) + 1);

    const tk = taxKey(row);
    if (tk !== "") taxSeen.set(tk, (taxSeen.get(tk) || 0) + 1);

    const ek = empIdKey(row);
    if (ek !== "") empIdSeen.set(ek, (empIdSeen.get(ek) || 0) + 1);

    const isActive = isBlank(row[retireIdx]);
    const personEmail = norm(row[emailColIdx]).toLowerCase();
    const manager = norm(row[mgrColIdx]) || "NoEmail"; // same key as grouping below
    if (!managerEmailCounts.has(manager)) {
      managerEmailCounts.set(manager, new Map<string, number>());
    }
    const mgrMapRef = managerEmailCounts.get(manager)!;
    if (isActive && personEmail && personEmail !== "null") {
      mgrMapRef.set(personEmail, (mgrMapRef.get(personEmail) || 0) + 1);
    }
  });

  const grouping = new Map<string, Record<string, Cell>[]>();

  // ---------- Per-row checks ----------
  outRows.forEach(row => {
    const managerEmail = norm(row[mgrColIdx]) || "NoEmail";
    const rec: Record<string, Cell> = {};

    const company = norm(row[companyColIndex]);
    const retireDate = row[retireIdx];
    const retireCause = row[colIndex("Retire Cause")];
    const isActive = isBlank(retireDate);

    // --- Cond7: Active flag must be consistent with Retire Date
    let activeValNum: number | null = null;
    {
      const activeRaw = row[colIndex("Active")];
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

    // --- Cond1: Retire Date / Retire Cause mismatch (both)
    const cond1 =
      (!isBlank(retireDate) && isBlank(retireCause)) ||
      (isBlank(retireDate) && !isBlank(retireCause));

    // --- Cond2: missing important field (actives only)
    const missingFields: string[] = importantCols.filter(c => isBlank(row[colIndex(c)]));
    const cond2 = isActive && missingFields.length > 0;

    // --- Cond3: duplicate Surname|Name|DOB within the same company (both)
    const cond3 = (nameDobSeen.get(nameDobKey(row)) || 0) > 1;

    const personEmail = norm(row[emailColIdx]).toLowerCase();
    const jobProperty = norm(row[colIndex("Job Property")]).toUpperCase();
    const isEmailValid = personEmail !== "" && personEmail !== "null";
    const mgrMapRef = managerEmailCounts.get(managerEmail);

    // --- Cond4: duplicate email under the same manager (actives only)
    const cond4 = isActive && isEmailValid && !!mgrMapRef && (mgrMapRef.get(personEmail) || 0) > 1;

    // --- Cond5: ADMINISTRATIVE without email (actives only)
    const cond5 = isActive && jobProperty === "ADMINISTRATIVE" && !isEmailValid;

    // --- Cond6: non-"alumil" email requires Mobile Phone (actives only)
    const mobilePhone = norm(row[colIndex("Mobile Phone")]);
    const nonAlumil = !personEmail.includes("alumil"); // blank counts as non-alumil
    const cond6 = isActive && nonAlumil && mobilePhone === "";

    // --- Cond8: company has no manager in Table3 (both)
    const cond8 = !mgrMap.has(company);

    // --- Cond9: duplicate Tax No. within the same company (both)
    const tk = taxKey(row);
    const cond9 = tk !== "" && (taxSeen.get(tk) || 0) > 1;

    // --- Cond10: duplicate Employee Id within the same company (both)
    const ek = empIdKey(row);
    const cond10 = ek !== "" && (empIdSeen.get(ek) || 0) > 1;

    // --- Date checks
    const hireS = toSerial(row[colIndex("Hire Date")]);
    const dobS = toSerial(row[colIndex("Date of Birth")]);
    const retireS = toSerial(retireDate);

    // Cond11: Retire Date earlier than Hire Date (leavers by nature)
    const cond11 = retireS !== null && hireS !== null && retireS < hireS;

    // Cond12: Hire Date or Date of Birth in the future (both)
    const hireFuture = hireS !== null && hireS > todaySerial;
    const dobFuture = dobS !== null && dobS > todaySerial;
    const cond12 = hireFuture || dobFuture;

    // Cond13: age under 16 at hire (both) or over 80 today (actives only)
    const underAgeAtHire = dobS !== null && hireS !== null && yearsBetween(dobS, hireS) < 16;
    const over80Today = isActive && dobS !== null && yearsBetween(dobS, todaySerial) > 80;
    const cond13 = underAgeAtHire || over80Today;

    // --- Supervisor checks (blank Supervisor Id is already covered by Cond2)
    const empId = norm(row[empIdIdx]);
    const supId = norm(row[supIdIdx]);
    const sameCompanyKey = company + "|" + supId;
    let supActive: boolean | undefined = undefined;
    if (supId !== "") {
      supActive = empActiveByCompanyId.has(sameCompanyKey)
        ? empActiveByCompanyId.get(sameCompanyKey)
        : empActiveById.get(supId);
    }

    // Cond14: Supervisor Id not found as an Employee Id (actives only)
    const cond14 = isActive && supId !== "" && supActive === undefined;

    // Cond15: active employee with a supervisor who is a leaver (actives only)
    const cond15 = isActive && supId !== "" && supActive === false;

    // Cond16: employee is their own supervisor (both)
    const cond16 = supId !== "" && empId !== "" && supId === empId;

    // --- Cond17: salary zero, negative or not a number (actives only; blanks covered by Cond2)
    const nominalRaw = row[colIndex("Nominal Salary")];
    const baseRaw = row[colIndex("Base Salary")];
    const nominalN = toNumber(nominalRaw);
    const baseN = toNumber(baseRaw);
    const badNominal = !isBlank(nominalRaw) && (isNaN(nominalN) || nominalN <= 0);
    const badBase = !isBlank(baseRaw) && (isNaN(baseN) || baseN <= 0);
    const cond17 = isActive && (badNominal || badBase);

    const shouldInclude =
      cond1 || cond2 || cond3 || cond4 || cond5 || cond6 || cond7 || cond8 || cond9 ||
      cond10 || cond11 || cond12 || cond13 || cond14 || cond15 || cond16 || cond17;

    rec["Cond1"] = cond1;
    rec["Cond2"] = cond2;
    rec["Cond3"] = cond3;
    rec["Cond4"] = cond4;
    rec["Cond5"] = cond5;
    rec["Cond6"] = cond6;
    rec["Cond7"] = cond7;
    rec["Cond8"] = cond8;
    rec["Cond9"] = cond9;
    rec["Cond10"] = cond10;
    rec["Cond11"] = cond11;
    rec["Cond12"] = cond12;
    rec["Cond13"] = cond13;
    rec["Cond14"] = cond14;
    rec["Cond15"] = cond15;
    rec["Cond16"] = cond16;
    rec["Cond17"] = cond17;
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
    if (cond8) reasons.push("Company has no manager in Table3");
    if (cond9) reasons.push("Duplicate Tax No. in same Company");
    if (cond10) reasons.push("Duplicate Employee Id in same Company");
    if (cond11) reasons.push("Retire Date before Hire Date");
    if (hireFuture) reasons.push("Hire Date in the future");
    if (dobFuture) reasons.push("Date of Birth in the future");
    if (underAgeAtHire) reasons.push("Under 16 at Hire Date");
    if (over80Today) reasons.push("Active employee over 80 years old");
    if (cond14) reasons.push("Supervisor Id not found");
    if (cond15) reasons.push("Supervisor is a leaver");
    if (cond16) reasons.push("Employee is own supervisor");
    if (badNominal && cond17) reasons.push("Nominal Salary zero, negative or invalid");
    if (badBase && cond17) reasons.push("Base Salary zero, negative or invalid");
    rec["Reason"] = reasons.join("; ");

    // Copy data columns (format dates)
    const dateColIndices: number[] = ["Hire Date", "Retire Date", "Date of Birth"].map(col => outHeaders.indexOf(col));
    outHeaders.forEach((h, i) => {
      const val = row[i];
      if (dateColIndices.includes(i) && typeof val === "number") {
        const jsDate = new Date(EXCEL_EPOCH + val * DAY_MS);
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
