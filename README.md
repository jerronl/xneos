# xneos

**xneos** is a lightweight integration tool that allows you to submit optimization jobs to the NEOS Server directly from Excel, powered by Python and [xlwings](https://www.xlwings.org/).

---

## ✨ Features

- Submit AMPL or other model jobs to NEOS directly from Excel
- Leverage Python via `xlwings` without leaving your spreadsheet
- Track job results and display outputs in Excel
- Simple setup with template-based project structure

---

## 🚀 Installation

### Option 1: From GitHub (latest)

```
pip install git+https://github.com/jerronl/xneos.git
```

### Option 2: From PyPI (when published)

```
pip install xneos
```

---

## 🧩 Usage

### 1. Install the xlwings Excel Add-in (Required)

```
xlwings addin install
```

This enables Excel to run Python functions via macros.

---

### 2. Create a New Project

Generate a new Excel+Python project using:

```
xneos quickstart myproject
```

This will create a folder `myproject/` with:
- `xneos_template.xlsm`: Excel file with a built-in "Solve" button (macro-enabled)
- `xneos_main.py`: Sample Python script for job submission
- `xlwings.conf`: Configuration for Python/Excel integration
- `manualstart.bat`: Manually start the python server

---

### 3. Run the Example

1. Open `xneos_template.xlsm`
2. Enable macros if prompted
3. Click the **Solve** button to submit a sample job to NEOS

---

### 4. Customize

You can update:
- The Excel sheet (e.g., input/output cells)
- The Python script (`xneos_main.py`) to match your own model and data
- The `.xlsm` macro to trigger additional logic

---

## ⚠️ Known Issues

### ❗ Run-time Error 1000: "No command specified in the configuration"

![Run-time error '1000'](images/error_1000.png)

This occurs when `xlwings` is unable to autostart the Python server due to misconfiguration.

### ✅ How to Fix:

- Make sure `xlwings` and its Excel Add-in are installed correctly
- Alternatively, start the UDF server manually:

```
myproject/manualstart.bat
```

Keep the terminal window open (you can minimize it).

---
## 📄 License

MIT License © 2025 Jerron
