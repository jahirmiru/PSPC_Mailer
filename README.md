# ✉️ PSPC Mailer (Plansculpt Auto Mailer Pro)

A modern, enterprise-grade desktop email automation suite built with Python and CustomTkinter. Designed specifically for academic outreach (MS/PhD applications, professor contact), research collaborations, and professional personalized bulk email campaigns.

---

## 🚀 Quick Start & Installation

### 🪟 For Windows Users (One-Click Launch)
If you are on Windows, simply double-click the **`run_app.bat`** file in the project folder. 
> 💡 **Tip:** It automatically checks for Python, installs any missing dependencies from `requirements.txt`, and launches the application without needing to type any commands!

---

### 💻 Manual Installation & Running (All Platforms)

Clone the repository and run the application using the following commands:

```bash
# 1. Clone the repository
git clone https://github.com/jahirmiru/PSPC_Mailer.git

# 2. Navigate to the project directory
cd PSPC_Mailer

# 3. Install required dependencies
pip install -r requirements.txt

# 4. Run the application
python pspc_mailer.py
```

---

## ✨ Key Features

- **🏢 Modern Dark & Light UI**: Built with CustomTkinter featuring clean typography and branding.
- **📝 VS Code Styled HTML Template Editor**:
  - Real-time syntax highlighting for HTML tags, attributes, strings, comments, and variables.
  - Interactive live visual preview with split-pane resizable divider.
  - Zoom controls (`A+`, `A-`, `100%`) and mouse wheel scaling.
- **⏰ 12-Hour AM/PM Interactive Campaign Scheduler**:
  - Easy dropdown picker for Target Date, Hours (`01-12`), Minutes (`00-55`), and `AM/PM`.
  - Multi-batch campaign queuing with automatic background execution.
  - Global timezone support.
- **🛡️ Smart Anti-Spam & Deliverability Engine**:
  - Customizable delays with random jitter (e.g., 10-25 seconds).
  - Batch cooldown pause (e.g., pause 5 minutes after every 20 emails).
  - Daily sending quotas to protect your email account reputation.
- **📊 Recipient & Data Management**:
  - Import CSV and Excel (`.xlsx`, `.xls`) files with interactive table preview.
  - Automatic duplicate removal and invalid row cleanup.
  - Dynamic placeholders (e.g., `{Professor_Name}`, `{Institute_Name}`, `{Recent_Paper_Topics}`, `{Email_Address}`).
- **📎 Multi-Attachment Support**: Easily attach CVs, cover letters, portfolios, or research papers.
- **⚙️ Universal SMTP Support**:
  - One-click presets for Gmail, Outlook / Office 365, Yahoo, and Custom SMTP servers.
  - Built-in "Test SMTP Connection" tool to verify credentials before launching campaigns.
- **📜 Live Activity Logs & CSV Export**: Real-time console logs and one-click detailed campaign delivery report export.

---

## 📦 Requirements

- **Python**: `3.9` or higher
- **Key Libraries**:
  - `customtkinter`
  - `pandas`
  - `openpyxl`
  - `tkinterweb`
  - `tkhtmlview`

*(All dependencies are listed in `requirements.txt`)*

---

## 🔒 Security & Best Practices

- **Gmail Users**: Use an **App Password** instead of your personal account password. (Enable 2-Step Verification in Google Account -> Security -> App Passwords).
- Sensitive credentials are not hardcoded and your session data remains strictly on your local machine.

---

## 📄 License
This project is open-source.
