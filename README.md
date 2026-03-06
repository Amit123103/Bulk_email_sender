# 📧 Bulk Email Pro v4.0
URL :::  https://bulk-email-sender-wsoy.onrender.com

**Send bulk emails from YOUR own email accounts.** Multi-account support, Excel upload, account rotation, spam score checker, A/B testing, scheduling, and full reports.

Three versions: **Desktop App** (Tkinter) • **Web App** (Flask) • **CLI Tool**

---

# 🚀 STEP-BY-STEP GUIDE: How to Send Emails Properly

Follow these steps **in order** to send your first email campaign successfully.

---

## Step 1: Install Dependencies

Open a terminal in the project folder and run:

```bash
cd "d:\Projects\automaticaliyemail sender"
pip install -r requirements.txt
```

If you don't have `requirements.txt`, install manually:
```bash
pip install pandas openpyxl xlrd cryptography keyring Pillow
```

---

## Step 2: Create an App Password (Gmail Example)

> ⚠️ **IMPORTANT:** You CANNOT use your regular Gmail password. You MUST create an App Password.

1. Go to **https://myaccount.google.com**
2. Click **"Security"** in the left menu
3. Scroll down and enable **"2-Step Verification"** (if not already enabled)
4. After enabling 2-Step Verification, go back to **Security**
5. Search for **"App Passwords"** (or go to https://myaccount.google.com/apppasswords)
6. Select App: **"Mail"**, Device: **"Windows Computer"**
7. Click **"Generate"**
8. **Copy the 16-character password** (looks like: `abcd efgh ijkl mnop`)
9. **Save this password** — you will need it in Step 4

### Other Providers

| Provider | SMTP Host | Port | Security | Password |
|----------|-----------|------|----------|----------|
| **Gmail** | `smtp.gmail.com` | 587 | TLS | App Password (16 chars) |
| **Outlook/Hotmail** | `smtp-mail.outlook.com` | 587 | TLS | Regular or App Password |
| **Yahoo** | `smtp.mail.yahoo.com` | 465 | SSL | App Password |
| **Zoho** | `smtp.zoho.com` | 587 | TLS | Regular password |
| **iCloud** | `smtp.mail.me.com` | 587 | TLS | App Password |
| **SendGrid** | `smtp.sendgrid.net` | 587 | TLS | API Key |
| **Custom/cPanel** | `mail.yourdomain.com` | 587 | TLS | Email password |

---

## Step 3: Launch the Desktop App

```bash
cd desktop_app
python bulk_email_pro.py
```

The app opens with 6 tabs: **Dashboard → Accounts → Recipients → Compose → Send → Reports**

---

## Step 4: Add Your Email Account

1. Click the **"📋 Accounts"** tab
2. Click **"+ Add Account"**
3. Fill in the form:
   - **Nickname:** Give it a name (e.g., "My Gmail")
   - **Email:** Your full email (e.g., `yourname@gmail.com`)
   - **Password:** Paste the **App Password** from Step 2
   - **Preset:** Select your provider (Gmail, Outlook, etc.) — it auto-fills SMTP settings
   - If no preset matches, manually enter:
     - **SMTP Host:** e.g., `smtp.gmail.com`
     - **Port:** e.g., `587`
     - **Security:** `TLS` (most common) or `SSL`
   - **Daily Limit:** 500 for Gmail, 300 for Outlook
4. Click **"🔌 Test Connection"** — must show ✅ **Connected**
5. Click **"Save"**

> 🔴 **If test fails with "getaddrinfo failed"**: Your SMTP host is wrong. Check the table above.
>
> 🔴 **If test fails with "Authentication failed"**: Your password is wrong. For Gmail, you MUST use the App Password, NOT your Google password.

---

## Step 5: Prepare Your Recipients Excel File

Create an Excel (.xlsx) or CSV file with your recipients:

| Email | Name | Company |
|-------|------|---------|
| john@example.com | John Smith | Acme Corp |
| jane@example.com | Jane Doe | Tech Inc |
| bob@example.com | Bob Wilson | StartupXYZ |

**Rules:**
- **One column MUST have email addresses** (column name should contain "email", "e-mail", or "mail")
- Other columns are optional — they become personalization variables
- Supported formats: `.xlsx`, `.xls`, `.csv`
- The system auto-detects the email column

---

## Step 6: Upload Recipients

1. Click the **"📂 Recipients"** tab
2. Click **"📂 Browse File"** and select your Excel/CSV file
3. The app auto-detects the email column
4. Click **"✅ Validate"**
5. Check the results:
   - ✅ Valid emails (will receive your email)
   - ❌ Invalid emails (bad format, removed)
   - 🔄 Duplicates (auto-removed)
   - 🔧 Typos fixed (e.g., `@gmail.con` → `@gmail.com`)

---

## Step 7: Compose Your Email

1. Click the **"✉️ Compose"** tab
2. **Select Send From:** Choose your email account from the dropdown
3. **Subject line:** Type your subject — use personalization like:
   ```
   Hello {{name}}, exciting news from {{company}}!
   ```
4. **Email body:** Write your email content. Use **HTML** for rich formatting:
   ```html
   <h2>Hello {{first_name}},</h2>
   <p>We have a great offer for <b>{{company}}</b>!</p>
   <p>Best regards,<br>Your Name</p>
   ```
5. **Use conditional blocks** (optional, advanced):
   ```
   {{#if company}}Dear {{company}} team,{{#else}}Dear Valued Customer,{{/else}}{{/if}}
   ```

### Available Personalization Variables

| Variable | Result |
|----------|--------|
| `{{name}}` | Full name from Excel |
| `{{first_name}}` | First name only |
| `{{last_name}}` | Last name only |
| `{{email}}` | Recipient email |
| `{{company}}` | Company name |
| `{{date}}` | Today's date (e.g., March 7, 2026) |
| `{{time}}` | Current time |
| `{{year}}` | Current year |
| `{{any_column}}` | ANY column from your Excel file |

### Compose Tab Buttons

| Button | What It Does |
|--------|-------------|
| **📄 Load Template** | Load a saved email template |
| **💾 Save Template** | Save current email as template |
| **🧪 Send Test** | Send one test email to yourself first |
| **👁 Preview** | Preview how the email looks |
| **🛡 Spam Check** | Analyze your email for spam triggers (score 0-100) |
| **📎 Attach File** | Attach a file to all emails |

---

## Step 8: Send a Test Email First (IMPORTANT!)

> ⚠️ **Always send a test email before launching a full campaign!**

1. In the Compose tab, click **"🧪 Send Test"**
2. Enter your own email address
3. Check your inbox — verify the email looks correct
4. Also click **"🛡 Spam Check"** to check your spam score (aim for < 30)

---

## Step 9: Launch the Campaign

1. Click the **"📤 Send"** tab
2. Configure settings:
   - **Delay:** 1.5-3.0 seconds between emails (recommended)
   - **Use Rotation:** ✅ Enable if you have multiple accounts
3. Click **"🚀 START SENDING"**
4. Watch the live progress:
   - Green = Sent successfully
   - Red = Failed (check error)
5. You can **Pause** or **Stop** anytime

---

## Step 10: View Reports

1. Click the **"📊 Reports"** tab after campaign finishes
2. See:
   - ✅ Total sent
   - ❌ Total failed
   - 📊 Success rate
3. Click **"Export Report"** to save as CSV

---

# 📋 QUICK REFERENCE: Common Errors & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| `getaddrinfo failed` | Wrong SMTP host | Check SMTP host in Step 4 table |
| `Authentication failed` | Wrong password | Use App Password, not regular password |
| `Connection timed out` | Firewall blocking port | Try port 465 with SSL, or check firewall |
| `Connection refused` | Wrong port | Check correct port in Step 4 table |
| `Daily limit reached` | Sent too many emails | Wait 24 hours, or add more accounts |
| `Too many login attempts` | Provider rate limit | Wait 15-30 minutes and try again |

---

# 🔧 Advanced Features

## Account Rotation
When **Use Account Rotation** is enabled:
- System cycles through all connected accounts automatically
- Switches when an account hits its daily send limit
- Gmail (500) + Outlook (300) + Yahoo (500) = **1,300 emails/day** from 3 accounts

## Spam Score Checker (🛡)
- Click **"🛡 Spam Check"** in Compose tab to analyze your email
- Checks for 50+ spam trigger words, phishing patterns, ALL CAPS, link quality
- Score 0-100: 🟢 0-30 = Good | 🟡 30-50 = Fair | 🟠 50-70 = Poor | 🔴 70+ = Spam

## A/B Split Testing
- Test different subject lines with split recipient groups
- Track which variant performs better
- Declare winner based on delivery rate

## Campaign Scheduling
- Schedule campaigns for a future date/time
- Support for recurring: daily, weekly, monthly

## Contact Segmentation
- Save validated recipients as named groups
- Tag contacts: VIP, New, Cold, Warm, Hot
- Filter by domain or tag, merge/split groups

## Template Library
- Save and load email templates
- Reuse your best-performing emails

## Auto-Save Draft
- Your compose state saves automatically when you close the app
- Restores when you reopen — never lose your work

---

# ⌨️ Keyboard Shortcuts (Desktop)

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Open Excel file |
| Ctrl+A | Add new email account |
| Ctrl+Enter | Start sending |
| Ctrl+P | Pause / Resume |
| Escape | Stop campaign |
| Ctrl+S | Save/Export report |

---

# 📁 Project Structure

```
desktop_app/                  ← Standalone Desktop Application
  bulk_email_pro.py           ← Main GUI app (Tkinter)
  account_manager.py          ← Account encryption & management
  excel_processor.py          ← Excel/CSV processing & validation
  smtp_engine.py              ← SMTP engine with rotation & anti-spam
  spam_checker.py             ← Email spam score analyzer
  scheduler.py                ← Campaign scheduling engine
  ab_tester.py                ← A/B split testing
  contact_manager.py          ← Contact groups & segmentation
  ai_generator.py             ← Optional AI email generation

web_app/                      ← Flask Web Application
  app.py                      ← Server with SocketIO + campaign history
  templates/                  ← HTML templates (dark theme)
  static/css/                 ← Responsive stylesheet

cli_tool/                     ← Command-Line Interface
  cli_sender.py               ← Click-based CLI with Rich output
```

---

# 🔒 Security

- **Passwords encrypted** with Fernet AES-128 (stored in `secret.key`)
- **Never stored in plain text** — encrypted in `accounts.json`
- Uses App Passwords (not your main email password)
- All connections use TLS/SSL encryption
- Account health monitoring with auto-disable

---

**Built with ❤️ — Bulk Email Pro v4.0**
