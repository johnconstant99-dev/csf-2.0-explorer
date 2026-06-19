# 🛡️ NIST CSF 2.0 Explorer

**Full-stack interactive reference and self-assessment tool** for the NIST Cybersecurity Framework 2.0.

[![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![Python](https://img.shields.io/badge/Python-3.12+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

---

## ✨ Features

| Feature                    | Description |
|---------------------------|-------------|
| **📊 Executive Dashboard**    | Live KPIs, progress by function, color-coded cards |
| **🔍 Browse & Search**        | Hierarchical or flat table view with powerful filtering |
| **📝 Self-Assessment**        | 5-level maturity scoring + notes per subcategory (persisted in SQLite) |
| **📈 Reports & Export**       | Progress charts, CSV export, auto-generated Markdown summary |
| **🏢 Organization Profile**   | Store org name and scope notes |

---

## 🚀 Quick Start (Local)

```bash
# 1. Clone the repo
git clone https://github.com/johnconstant99-dev/csf-2.0-explorer.git
cd csf-2.0-explorer

# 2. Install dependencies
pip install -r requirements.txt

# 3. (Optional but recommended) Download the latest CSF 2.0 Core export
#    from the official NIST Reference Tool and place it as:
#    cprt_CSF_2_0_0_06-17-2026.xlsx  (in the same folder)

# 4. Run the app
streamlit run csf_2_0_explorer.py
```

The app will open in your browser at `http://localhost:8501`.

---

## 📦 What's Included

- `csf_2_0_explorer.py` — Complete Streamlit application
- `requirements.txt` — Minimal dependencies
- `README.md` — This file

**Data file** (`cprt_CSF_2_0_0_06-17-2026.xlsx`):  
Not included in the repo (to keep it lightweight).  
Export it yourself from the [NIST CSF 2.0 Reference Tool](https://csrc.nist.gov/projects/cybersecurity-framework/csf-2-0-reference-tool) or use any recent Core export. The parser is robust and will work with the standard format.

---

## 🗂️ Project Structure

```
csf-2.0-explorer/
├── csf_2_0_explorer.py      # Main application
├── requirements.txt
├── README.md
└── csf_assessments.db       # Created automatically on first run (gitignored)
```

---

## 🛠️ Tech Stack

- **Frontend**: Streamlit (reactive components, data editor, charts)
- **Backend**: Python 3.12 + pandas + openpyxl
- **Database**: SQLite (zero-config, local persistence)
- **Data Source**: NIST CSF 2.0 Core (June 2026 export)

---

## 📝 Data Notes

- Parses **185 subcategories** across all 6 Functions
- 79 subcategories are marked as **Withdrawn** (moved/consolidated in v2.0) — hidden by default
- Implementation Examples are fully included
- Informative References are empty in most exports (official mappings live on nist.gov)

---

## ⚠️ Disclaimer

This is an **unofficial** community tool created for reference, education, and internal self-assessment purposes.  
It is **not affiliated with or endorsed by NIST**.

For official documentation and the latest CSF 2.0 materials, visit:  
https://www.nist.gov/cyberframework

Feedback on the framework itself should be sent to: **cprt@nist.gov**

---

## 🤝 Contributing

Pull requests and issues are welcome! Feel free to suggest improvements to the UI, add new assessment features, or enhance the reporting capabilities.

---

Built with ❤️ for cybersecurity, GRC, and risk management professionals.
