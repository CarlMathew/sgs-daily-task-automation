# SGS Daily Automation Task

This project automates daily reporting tasks for SGS sites using Python scripts and Excel macros.

---

## ⚙️ Setup Instructions

### 1. Create a Python Virtual Environment

```bash
python -m venv venv
```

Activate the environment:

- **Windows:**

  ```bash
  .\venv\Scripts\activate
  ```

- **macOS/Linux:**

  ```bash
  source venv/bin/activate
  ```

---

### 2. Install Dependencies

Make sure you're inside the virtual environment, then run:

```bash
pip install -r requirements.txt
```

---

### 3. Configure Excel Macros

In the `macros/` folder, you will find:

- `Copy of PERSONAL.xlsb`
- `DailyTaskFixFormat.xlsm`

Move both files to your **Documents** folder:

```
Documents/
├── PERSONAL.xlsb
└── DailyTaskFixFormat.xlsm
```

These macros help automate Excel formatting for daily tasks.

---

### 4. Folder Structure

Each site should have a `report/` subfolder inside its respective folder under `data/`.

Example:

```
data/
├── Scott/
│   └── report/
├── Texas/
│   └── report/
```

---

### 5. Running the Scripts

From the project root directory, run:

- **Worklist Script:**

  ```bash
  python worklist.py
  ```

- **Rush Script:**

  ```bash
  python rush.py
  ```

---

## 📁 Example Project Structure

```
SGS-Automation/
├── data/
│   ├── Scott/
│   │   └── report/
│   └── Texas/
│       └── report/
├── macros/
│   ├── Copy of PERSONAL.xlsb
│   └── DailyTaskFixFormat.xlsm
├── requirements.txt
├── worklist.py
└── rush.py
```

---

## 📝 Notes

- Enable macros in Excel for full functionality.
- If errors occur, check terminal output or log files.
- Make sure all folder paths and files are correctly placed.
