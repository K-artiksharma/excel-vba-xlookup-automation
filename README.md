# Excel VBA XLOOKUP Automation via Multi-Input Search

This project showcases a dynamic Excel VBA automation tool that performs **XLOOKUP-style record retrieval** from a source sheet based on **user input through an InputBox**. Users can input one or more **IDs or Names**, separated by commas, and the tool fetches all matching rows and transfers them to the destination sheet.

---

## 💡 How It Works

1. The user clicks the **XLOOKUP button** on the `XLOOKUP` sheet.
2. An **InputBox** appears prompting the user to enter one or more **IDs or Names** (comma-separated).
3. The VBA script:
   - Splits the input into individual search terms.
   - Searches each value in the `Restaurant` sheet (source).
   - Retrieves the corresponding row(s).
   - Pastes them sequentially into the **next available empty rows** in the `XLOOKUP` sheet (destination).

---

## ✅ Features

- Supports **multiple value lookups** in a single input (e.g., `101, 205, David`).
- Works with **ID or Name** as lookup keys.
- Auto-finds and pastes results into the next available empty rows.
- **XLOOKUP-style logic using VBA**, eliminating manual formulas.
- Interactive and user-friendly with input prompts.
- Fully documented and modular VBA code for easy maintenance.

---

## 📁 File Contents

- `XLOOKUP_VBA_automation.xlsm` – Excel macro-enabled workbook containing:
  - `Restaurant` – Source data sheet (e.g., ID, Name, Country, City, Rating, etc.)
  - `XLOOKUP` – Destination sheet where results are displayed
  - VBA Module – Contains button-triggered script logic

---

## 🚀 Use Cases

- Fast, repeated data retrieval from large tables
- CRM/Restaurant/order management tools
- Automating manual lookup and copy-paste tasks
- Dynamic search operations during data entry or review

---

## 🧑‍💻 Technologies Used

- Microsoft Excel (.xlsm)
- VBA (Visual Basic for Applications)
- InputBox, Loops, String Manipulation, Range Search

---

## 📝 How to Use

1. Download and open the `.xlsm` file.
2. Enable Macros when prompted.
3. Go to the `XLOOKUP` sheet.
4. Click the **XLOOKUP button**.
5. In the InputBox, enter one or more **IDs or Names**, separated by commas (e.g., `102, Amit, 205`).
6. Matching rows from the `Restaurant` sheet will be copied to the next empty rows of the `XLOOKUP` sheet.

---
