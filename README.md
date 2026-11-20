# 📋 NC Registration System — Web App

A full-featured Non-Conformance (NC) registration and tracking web application built with **Streamlit**, converted from the original Tkinter desktop app.

---

## ✨ Features

| Feature | Detail |
|---|---|
| Auto NC Number | Generated as `NC-2026-001`, `NC-2026-002`, … (duplicate-safe) |
| Auto Date / Month / Week | Filled automatically from the server clock |
| Customer dropdown | Bosch, Continental, Valeo, Cummins, Delphi |
| Function dropdown | SMT, QA, Production, Maintenance, Engineering, Stores |
| Auto Responsibility | Assigned from Function using fixed mapping |
| Stage dropdown | Incoming, SPI, Placement, Reflow, AOI, Testing, Packing |
| NC Category | Major / Minor |
| 6M Category | Man, Machine, Method, Material, Measurement, Mother Nature |
| Validation | Blocks save when required fields are empty |
| Excel storage | Records saved to `NC_Registration.xlsx` (GitHub-backed) |
| Excel download | Download all records directly from the browser |
| Persistent storage | Records survive app restarts via GitHub repository |

---

## 🗂️ Project Structure

```
.
├── app.py              ← Streamlit web application (main entry point)
├── final code.py       ← Original Tkinter desktop app (untouched)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
└── NC_Registration.xlsx← Auto-created when first record is saved
```

---

## 🖥️ Run Locally

### 1 — Install dependencies
```bash
pip install -r requirements.txt
```

### 2 — Start the app
```bash
streamlit run app.py
```

The app opens automatically at **http://localhost:8501** in your browser.

> **Local mode:** Records are saved to `NC_Registration.xlsx` in the project folder.
> No GitHub secrets are required for local use.

---

## ☁️ Deploy on Streamlit Community Cloud

### Step 1 — Push to GitHub

1. Create a **public** (or private) GitHub repository named e.g. `nc-registration`.
2. Open a terminal in the project folder and run:

```bash
git init                          # (skip if .git already exists)
git add app.py requirements.txt README.md
git commit -m "Initial commit – NC Registration Streamlit app"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/nc-registration.git
git push -u origin main
```

> ⚠️ Do **not** commit `NC_Registration.xlsx` or any secrets.
> Add them to `.gitignore`:
> ```
> NC_Registration.xlsx
> .streamlit/secrets.toml
> ```

---

### Step 2 — Create a GitHub Personal Access Token (PAT)

The app uses the GitHub API to persist `NC_Registration.xlsx` back into your repo
so records are never lost between Streamlit Cloud restarts.

1. Go to **GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)**.
2. Click **Generate new token (classic)**.
3. Give it a name like `nc-streamlit`.
4. Select the **`repo`** scope (full control of private repositories).
5. Click **Generate token**.
6. **Copy the token** — you will not see it again.

---

### Step 3 — Deploy on Streamlit Community Cloud

1. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub.
2. Click **New app**.
3. Select:
   - **Repository:** `<YOUR_USERNAME>/nc-registration`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Advanced settings → Secrets** and paste:

```toml
GITHUB_TOKEN     = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
GITHUB_REPO      = "YOUR_USERNAME/nc-registration"
GITHUB_FILE_PATH = "NC_Registration.xlsx"
```

5. Click **Deploy**.

Your app will be live at:
```
https://<your-app-name>.streamlit.app
```

---

### Step 4 — Verify persistent storage

1. Open the app URL in Chrome.
2. Fill in a registration and click **Save Registration**.
3. Refresh the page — the record should still appear.
4. Go to your GitHub repo — you should see `NC_Registration.xlsx` committed there.

---

## 🔐 Secrets Reference

| Secret key | Description |
|---|---|
| `GITHUB_TOKEN` | Personal Access Token with `repo` scope |
| `GITHUB_REPO` | `username/repository-name` |
| `GITHUB_FILE_PATH` | Path inside repo, default `NC_Registration.xlsx` |

> 🔒 Secrets are stored encrypted by Streamlit and are never exposed to users.

---

## 📋 Responsibility Mapping

| Function | Responsible Person |
|---|---|
| SMT | Santosh |
| QA | Vidhyashree |
| Production | Karthik |
| Maintenance | Basavaraj |
| Engineering | Sujith |
| Stores | Prashanth |

---

## 🛠️ Tech Stack

- **[Streamlit](https://streamlit.io)** – Web framework
- **[Pandas](https://pandas.pydata.org)** – Data manipulation
- **[OpenPyXL](https://openpyxl.readthedocs.io)** – Excel read/write
- **[Requests](https://docs.python-requests.org)** – GitHub API calls

---

## 📄 License

Internal use only. Original desktop application by the project team.
