# Goal Progress Dashboard

A live Flask dashboard that reads your Excel file and visualizes goal progress — with automatic change detection and highlighting.

---

## Requirements

- Python 3.8 or higher
- The Excel file: `Goal_progress_June.xlsx`

Check your Python version:
```bash
python3 --version
```

---

## Quick Start (Flask already installed)

```bash
cd /path/to/keerthi
python3 app.py
```

Then open your browser at: **http://localhost:8080**

---

## Setup on a New Machine (Flask not installed)

### Step 1 — Clone or copy the project folder

Make sure you have these files:
```
keerthi/
├── app.py
├── requirements.txt
├── README.md
└── templates/
    └── dashboard.html
```

### Step 2 — Install dependencies

```bash
pip3 install -r requirements.txt
```

This installs:
- `flask` — the web server
- `openpyxl` — reads `.xlsx` files

> If `pip3` is not found, try `pip` instead.

### Step 3 — Point to your Excel file

Open `app.py` and update line 10 to the path of your Excel file:

```python
EXCEL_FILE = "/path/to/your/Goal_progress_June.xlsx"
```

### Step 4 — Run the dashboard

```bash
python3 app.py
```

Open your browser at: **http://localhost:8080**

---

## Troubleshooting

### Port already in use
macOS uses port 5000 for AirPlay. This app runs on **8080** to avoid that.
If 8080 is also taken, change the port in `app.py` line 191:
```python
app.run(debug=True, port=9000, use_reloader=False)
```
Then open **http://localhost:9000**

### pip3 not found
```bash
python3 -m pip install -r requirements.txt
```

### Python not found
Download from: https://www.python.org/downloads/

### Excel file not found error
The app will show a 500 error if the Excel file path is wrong.
Double-check the `EXCEL_FILE` path in `app.py` line 10.

---

## How Change Detection Works

1. Every **30 seconds** the browser polls `/api/data`
2. Flask checks the **MD5 hash** of the Excel file
3. If the hash changed → it re-reads the file and diffs every cell
4. Changed cells are highlighted in **amber** on the charts
5. A `🔔 N updates` badge appears in the top-right header
6. Click the badge to open the **Change Log** panel showing before/after values

To test it: open the Excel file, change any cell, save — the dashboard will update within 30 seconds.

---

## Stopping the Server

Press `Ctrl + C` in the terminal where the server is running.
