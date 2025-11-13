# 📊 Universal Excel Automation — *Ultimate Auto Mode*

**A fully automated, intelligent Excel analysis engine that works with *any* Excel file — from *any* department — with *zero configuration*.**

This project automatically:

* Detects **numeric**, **categorical**, **date**, and **ID-like** columns
* Cleans and prepares the dataset
* Generates multi-sheet Excel reports
* Performs summary statistics
* Identifies outliers
* Produces monthly trends (if date columns exist)
* Analyzes missing values
* Extracts top categorical values
* Works with *any* Excel schema — students, HR, finance, sales, logistics, anything.

---

# 🚀 Features

### ✅ **1. Automatic Column Type Detection**

The engine intelligently classifies each column as:

* **Numeric**
* **Categorical**
* **Date**
* **Possible Unique Identifier**

### ✅ **2. Universal Support for Any Excel Dataset**

Works even if column names are:

* Unknown
* Different across departments
* Messy
* In random order

### Example:

* Student dataset → Marks, Attendance, City
* HR dataset → Salary, Department, DOJ
* Sales dataset → Invoice Date, Amount, Item, Region
* Logistics dataset → Route, Cost, Delivery Date

All handled automatically.

### ✅ **3. Smart Summaries Generated Automatically**

* **Numeric summary** (sum, mean, median, std, min, max)
* **Categorical summary** (top 10 values)
* **Date-based summary** (month-wise trend)
* **Missing value summary**
* **Outlier detection (IQR method)**
* **ID candidate detection** (unique, high-cardinality columns)
* **Top rows for each numeric column**

---

# 📁 Project Structure

```
universal_excel_automation/
│
├── data/                     # Put your raw Excel files here
│   ├── sample_sales.xlsx
│   └── sample_students.xlsx
│
├── reports/                  # Auto-generated reports appear here
│   ├── sample_sales_report.xlsx
│   └── sample_students_report.xlsx
│
├── src/
│   └── generate_universal_report.py  # Main engine
│
├── README.md
└── EXPLANATION.md
```

---

# 🛠️ Installation

### **1. Clone the repository**

```bash
git clone https://github.com/0723naman/universal-excel-automation.git
cd universal-excel-automation
```

### **2. Create virtual environment**

**Windows PowerShell**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**macOS/Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### **3. Install required packages**

```bash
pip install pandas openpyxl
```

---

# ▶️ Usage

### **Option 1 — Process all Excel files inside `/data`**

```bash
python src/generate_universal_report.py
```

### **Option 2 — Process a specific Excel file**

```bash
python src/generate_universal_report.py --input data/sample_sales.xlsx
```

Reports will appear in the `/reports` folder automatically.

---

# 📈 Example Output Sheets

Each generated report includes:

### 🔹 RawData

Original file (cleaned formatting only)

### 🔹 NumericSummary

* Sum, Mean, Median, Std, Min, Max
* Missing value count

### 🔹 Categorical “Top Values” Sheets

One sheet per categorical column (top 10 values)

### 🔹 Monthly Sheets

(If a date column exists)

* Month-wise aggregation of numeric data
* Row count trends

### 🔹 Outliers

Detected using IQR method

### 🔹 MissingValues

Count of missing entries per column

### 🔹 ID_Candidates

Columns likely representing unique identifiers

### 🔹 TopRows per Numeric Column

Top 10 rows sorted by each numeric column

---

# 🧠 How It Works Internally (Short)

1. Columns → automatically classified
2. Dates → converted safely
3. Numeric columns → summarized
4. Categorical → frequency analysis
5. Date columns → monthly trends
6. Outliers → detected using IQR
7. Missing values → counted
8. ID candidates → selected via uniqueness
9. All outputs saved as sheets in a single Excel report

---

# 🎯 Ideal Use Cases

* Finance teams (monthly sales, expenses, KPIs)
* HR analytics (salary, attendance, joining trends)
* School/college data (marks, attendance, admissions)
* Marketing (campaign performance)
* Logistics (delivery date, cost trends)
* Operations dashboards
* Business intelligence preprocessing

---

# 🎁 Sample Files Included

* `sample_sales.xlsx`
* `sample_students.xlsx`

Use them to test the pipeline.

---

# 🤝 Contributing

Pull requests are welcome.
For major changes, open an issue first to discuss your idea.

---

# 📜 License

MIT License (can be changed if you prefer another)

---

# ⭐ If you found this useful

Please ⭐ star the repository — it really helps!
