# Handsoff Automation Studio 🚀
### Multi-User Bank Statement ➔ Manager Verification ➔ Handsoff Automation System

---

## 🔒 Manager Data Privacy & Security (Crucial Feature)

Portal me **strict role-based privacy** implement ki gayi hai:

1. **Manager View (e.g. Smita, Ambuj, Aditi, etc.):**
   - **Blank Data (Unallocated Pool):** Manager ko sirf unallocated transactions dikhti hain taaki wo customer payment screenshot (Amount / UTR) dekhkar use apni team ke naam par claim kar sake.
   - **Apni Team Ka Data:** Manager ko sirf wo data dikhega jo uski apni team ne claim kiya hai.
   - **Baki Managers Ka Data:** **100% Invisible & Hidden**! Database query level par hi filter ho jata hai, doosre managers ka data kisi ko show nahi hota.

2. **Admin / Backend View:**
   - Admin sabhi bank statements upload kar sakta hai.
   - Sabhi managers ki allocations aur pending unallocated pool ka live summary dekh sakta hai.
   - 1-Click me Master **Handsoff Format (.xlsx)** aur **Consolidated Paid File (.xlsx)** export kar sakta hai.

---

## 📂 System Files

- [**`start.bat`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/start.bat) — Double click to launch the portal on `http://localhost:8085`.
- [**`server.py`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/server.py) — Web server & APIs with port auto-free.
- [**`db_manager.py`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/db_manager.py) — SQLite database engine enforcing strict manager privacy.
- [**`core_engine.py`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/core_engine.py) — Excel parsing, UTR extraction, formulas (`=TEXT(...)`, `=N&O`).
- [**`database/handsoff.db`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/database/handsoff.db) — Persistent SQLite database.
- [**`templates/index.html`**](file:///d:/New%20folder/New%20folder/Handsoff%20Automation/templates/index.html) — Modern dashboard UI with active role switcher.

---

## 🚀 Daily Workflow (Agle Din Ka Process)

### Step 1: Subah Admin Bank Statement Upload Karega
1. Admin portal open karega aur **"👑 Admin / Backend"** select karega.
2. **"Admin Statement Upload"** tab me daily `Account Transaction Format.xlsx` upload karega.
3. System bina kisi purane data ko delete kiye nayi transactions database me sync kar dega.

### Step 2: Managers Apni Payment Verify & Claim Karenge
1. Manager apna naam select karega (e.g. `💼 Smita Jagtap`).
2. Screen par sirf **Blank (Unallocated) Bank Transactions** dikhenge.
3. Customer ke payment screenshot se **Amount** ya **UTR** search box me type karega.
4. Transaction par **"Claim / Update"** click karke sirf 3 details bharega:
   - **Customer ID / Loan No**
   - **Customer Name**
   - **Caller Name** (Portfolio auto-mapped)
5. Save karte hi wo entry us manager ke naam par allocate ho jayegi. Ab koi doosra manager use claim nahi kar sakta aur na hi dekh sakta hai!

### Step 3: Backend Team Master Handsoff Download Karegi
1. Admin / Backend **"Handsoff (.xlsx)"** button dabayega.
2. Exact standard format me:
   - Bank Reference No (`H261...`, `I307...`, `R10...`)
   - Extracted Cheque/UTR No
   - Day Formula: `=TEXT(H{row},"DDDDDD")`
   - Concate Formula: `=N{row}&O{row}`
   - Portfolio Source & Product Type
   ke sath final Excel download ho jayegi!
