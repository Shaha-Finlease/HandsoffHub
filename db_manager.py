import os
import sqlite3
import datetime
import core_engine

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, 'database')
os.makedirs(DB_DIR, exist_ok=True)
DB_PATH = os.path.join(DB_DIR, 'handsoff.db')

DEFAULT_MANAGERS = [
    {"name": "AMBUJ TRIPATHI", "agency": "INH"},
    {"name": "SMITA JAGTAP", "agency": "INH"},
    {"name": "ADITI", "agency": "INH"},
    {"name": "OMKAR SINGH", "agency": "INH"},
    {"name": "SHRAVANTI", "agency": "INH"},
    {"name": "SHWETA", "agency": "INH"},
    {"name": "SATISH SIR", "agency": "INH"},
    {"name": "POONAM INDULKAR", "agency": "KP"}
]

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    c = conn.cursor()
    
    # Bank Transactions Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS bank_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ref_no TEXT,
            bank_name TEXT,
            sheet_name TEXT,
            tx_date TEXT,
            amount REAL,
            narration TEXT,
            cheque_raw TEXT,
            cheque_extracted TEXT,
            batch_date TEXT,
            status TEXT DEFAULT 'UNALLOCATED',
            claimed_by_manager TEXT,
            customer_id TEXT,
            customer_name TEXT,
            portfolio TEXT,
            caller_name TEXT,
            mode_of_pay TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(ref_no, bank_name, amount)
        )
    ''')
    
    # Managers Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS managers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            agency TEXT DEFAULT 'INH'
        )
    ''')
    
    # Insert default managers
    for m in DEFAULT_MANAGERS:
        c.execute('''
            INSERT OR IGNORE INTO managers (name, agency)
            VALUES (?, ?)
        ''', (m["name"], m["agency"]))
        
    conn.commit()
    conn.close()

def insert_transactions(tx_list, batch_date=None):
    if not batch_date:
        batch_date = datetime.datetime.now().strftime("%d-%m-%Y")
        
    conn = get_connection()
    c = conn.cursor()
    inserted = 0
    duplicates = 0
    
    for tx in tx_list:
        ref_no = tx.get("ref") or f"TX-{int(datetime.datetime.now().timestamp()*1000)}"
        bank_name = tx.get("bank", "BANK")
        sheet_name = tx.get("sheet", "")
        tx_date = tx.get("date", "")
        amount = float(tx.get("amount", 0.0))
        narration = tx.get("narration", "")
        cheque_raw = tx.get("cheque_ref", "")
        cheque_extracted = core_engine.extract_cheque_no(narration, cheque_raw)
        
        try:
            c.execute('''
                INSERT INTO bank_transactions (
                    ref_no, bank_name, sheet_name, tx_date, amount, narration,
                    cheque_raw, cheque_extracted, batch_date, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'UNALLOCATED')
            ''', (ref_no, bank_name, sheet_name, tx_date, amount, narration, cheque_raw, cheque_extracted, batch_date))
            inserted += 1
        except sqlite3.IntegrityError:
            duplicates += 1
            
    conn.commit()
    conn.close()
    return inserted, duplicates

def get_transactions(role, manager_name=None, status_filter=None, search_query=None, bank_filter=None, amount_filter=None, date_filter=None):
    """
    STRICT PRIVACY with SMART MULTI-PARAM SEARCH:
    - If role == 'ADMIN': sees everything.
    - If role == 'MANAGER': sees ONLY 'UNALLOCATED' (blank data) OR data claimed by manager_name!
    - Multi-token keyword search: matches Amount, UTR (full or partial), Customer Name, Bank, Narration keywords.
    """
    conn = get_connection()
    c = conn.cursor()
    
    query = "SELECT * FROM bank_transactions WHERE 1=1"
    params = []
    
    if role != "ADMIN":
        query += " AND (status = 'UNALLOCATED' OR UPPER(TRIM(claimed_by_manager)) = UPPER(TRIM(?)))"
        params.append(manager_name or "")
        
    if status_filter:
        query += " AND status = ?"
        params.append(status_filter)

    if bank_filter and bank_filter.strip():
        query += " AND UPPER(bank_name) = UPPER(?)"
        params.append(bank_filter.strip())

    if date_filter and date_filter.strip():
        query += " AND (tx_date LIKE ? OR batch_date LIKE ?)"
        d_val = f"%{date_filter.strip()}%"
        params.extend([d_val, d_val])

    if amount_filter:
        try:
            amt_num = float(str(amount_filter).replace(',', '').strip())
            # Match within 1.0 tolerance (e.g. rounding or minor fee difference)
            query += " AND ABS(amount - ?) < 1.0"
            params.append(amt_num)
        except Exception:
            pass
        
    if search_query:
        # Split search query into multiple keywords / tokens
        tokens = [t.strip() for t in search_query.split() if t.strip()]
        for token in tokens:
            t_like = f"%{token}%"
            query += """ AND (
                narration LIKE ? OR 
                ref_no LIKE ? OR 
                cheque_extracted LIKE ? OR 
                cheque_raw LIKE ? OR
                customer_name LIKE ? OR 
                customer_id LIKE ? OR 
                caller_name LIKE ? OR 
                bank_name LIKE ? OR
                CAST(amount AS TEXT) LIKE ?
            )"""
            params.extend([t_like, t_like, t_like, t_like, t_like, t_like, t_like, t_like, t_like])
        
    query += " ORDER BY id DESC"
    
    c.execute(query, params)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows

def smart_match_finder(role, manager_name, amount=None, partial_name=None, partial_utr=None, bank=None):
    """
    Heuristic / Fuzzy finder when UTR is wrong or missing.
    Ranks unallocated candidates based on matching signals.
    """
    # Fetch all candidate transactions visible to user
    candidates = get_transactions(role, manager_name, status_filter='UNALLOCATED', bank_filter=bank)
    ranked = []
    
    amt_target = None
    if amount is not None:
        try:
            amt_target = float(str(amount).replace(',', '').strip())
        except Exception:
            pass

    p_name = (partial_name or "").lower().strip()
    p_utr = (partial_utr or "").lower().strip()

    for tx in candidates:
        score = 0
        reasons = []
        
        tx_amt = float(tx["amount"])
        tx_narr = (tx["narration"] or "").lower()
        tx_cheque = (tx["cheque_extracted"] or "").lower()
        tx_raw_cheque = (tx["cheque_raw"] or "").lower()

        # 1. Amount match
        if amt_target is not None:
            if abs(tx_amt - amt_target) < 0.01:
                score += 50
                reasons.append("Exact Amount Match (Rs. " + str(tx_amt) + ")")
            elif abs(tx_amt - amt_target) <= 10.0:
                score += 20
                reasons.append(f"Near Amount (Diff: Rs. {round(abs(tx_amt - amt_target), 2)})")

        # 2. UTR partial match (even if only 4-6 digits match)
        if p_utr:
            if p_utr in tx_cheque or p_utr in tx_raw_cheque or p_utr in tx_narr:
                score += 40
                reasons.append(f"Partial UTR Match ('{p_utr}')")
            else:
                # Check last 4 or 6 digits
                clean_utr = "".join(filter(str.isalnum, p_utr))
                if len(clean_utr) >= 4:
                    last4 = clean_utr[-4:]
                    if last4 in tx_cheque or last4 in tx_narr:
                        score += 30
                        reasons.append(f"Last 4 digits UTR Match ('{last4}')")

        # 3. Customer name match in narration
        if p_name and len(p_name) >= 3:
            # Check individual name parts (e.g. 'Saravanan' or 'Vetry')
            for part in p_name.split():
                if len(part) >= 3 and part in tx_narr:
                    score += 35
                    reasons.append(f"Name Match in Narration ('{part}')")

        if score > 0:
            tx_copy = dict(tx)
            tx_copy["match_score"] = score
            tx_copy["match_reasons"] = reasons
            ranked.append(tx_copy)

    # Sort descending by score
    ranked.sort(key=lambda x: x["match_score"], reverse=True)
    return ranked

def claim_transaction(tx_id, manager_name, customer_id, customer_name, portfolio, caller_name, mode_of_pay=None):
    conn = get_connection()
    c = conn.cursor()
    
    # Check if already claimed by someone else
    c.execute("SELECT status, claimed_by_manager, narration, cheque_raw FROM bank_transactions WHERE id = ?", (tx_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return False, "Transaction not found."
        
    if row["status"] == "ALLOCATED" and row["claimed_by_manager"] and row["claimed_by_manager"].upper().strip() != manager_name.upper().strip():
        conn.close()
        return False, f"Already claimed by Manager: {row['claimed_by_manager']}"
        
    # Auto-extract mode and cheque if missing
    narration = row["narration"] or ""
    if not mode_of_pay:
        mode_of_pay = core_engine.detect_mode("", narration)
    cheque_ext = core_engine.extract_cheque_no(narration, row["cheque_raw"])
    
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''
        UPDATE bank_transactions SET
            customer_id = ?,
            customer_name = ?,
            portfolio = ?,
            caller_name = ?,
            claimed_by_manager = ?,
            mode_of_pay = ?,
            cheque_extracted = ?,
            status = 'ALLOCATED',
            updated_at = ?
        WHERE id = ?
    ''', (customer_id, customer_name, portfolio, caller_name, manager_name, mode_of_pay, cheque_ext, now_str, tx_id))
    
    conn.commit()
    conn.close()
    return True, "Successfully updated and claimed!"

def unclaim_transaction(tx_id, manager_name, role):
    conn = get_connection()
    c = conn.cursor()
    
    c.execute("SELECT claimed_by_manager FROM bank_transactions WHERE id = ?", (tx_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return False, "Transaction not found."
        
    if role != "ADMIN" and row["claimed_by_manager"] and row["claimed_by_manager"].upper().strip() != manager_name.upper().strip():
        conn.close()
        return False, "You can only unclaim your own team's transactions."
        
    c.execute('''
        UPDATE bank_transactions SET
            customer_id = NULL,
            customer_name = NULL,
            portfolio = NULL,
            caller_name = NULL,
            claimed_by_manager = NULL,
            status = 'UNALLOCATED',
            updated_at = ?
        WHERE id = ?
    ''', (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tx_id))
    
    conn.commit()
    conn.close()
    return True, "Transaction released back to blank / unallocated pool."

def get_stats(role, manager_name=None):
    conn = get_connection()
    c = conn.cursor()
    
    stats = {}
    
    # Unallocated
    c.execute("SELECT COUNT(*), COALESCE(SUM(amount), 0) FROM bank_transactions WHERE status = 'UNALLOCATED'")
    unalloc_row = c.fetchone()
    stats["unallocated_count"] = unalloc_row[0]
    stats["unallocated_amount"] = round(unalloc_row[1], 2)
    
    if role == "ADMIN":
        # Admin sees full overview
        c.execute("SELECT COUNT(*), COALESCE(SUM(amount), 0) FROM bank_transactions")
        tot_row = c.fetchone()
        stats["total_count"] = tot_row[0]
        stats["total_amount"] = round(tot_row[1], 2)
        
        c.execute("SELECT COUNT(*), COALESCE(SUM(amount), 0) FROM bank_transactions WHERE status = 'ALLOCATED'")
        alloc_row = c.fetchone()
        stats["allocated_count"] = alloc_row[0]
        stats["allocated_amount"] = round(alloc_row[1], 2)
        
        # Breakdown by Manager
        c.execute('''
            SELECT claimed_by_manager, COUNT(*), COALESCE(SUM(amount), 0)
            FROM bank_transactions
            WHERE status = 'ALLOCATED' AND claimed_by_manager IS NOT NULL
            GROUP BY claimed_by_manager
        ''')
        stats["by_manager"] = [
            {"manager": r[0], "count": r[1], "amount": round(r[2], 2)}
            for r in c.fetchall()
        ]
    else:
        # Manager sees ONLY their team stats + unallocated pool count
        m_norm = (manager_name or "").upper().strip()
        c.execute('''
            SELECT COUNT(*), COALESCE(SUM(amount), 0)
            FROM bank_transactions
            WHERE status = 'ALLOCATED' AND UPPER(TRIM(claimed_by_manager)) = ?
        ''', (m_norm,))
        my_row = c.fetchone()
        stats["my_allocated_count"] = my_row[0]
        stats["my_allocated_amount"] = round(my_row[1], 2)
        
    conn.close()
    return stats

def get_managers_list():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT name, agency FROM managers ORDER BY name ASC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows

# Initialize DB on load
init_db()
