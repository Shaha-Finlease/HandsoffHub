import os
import sys
import json
import time
import subprocess
import webbrowser
from threading import Timer
from flask import Flask, request, jsonify, send_file, render_template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)
sys.path.append(BASE_DIR)

import core_engine
import db_manager

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'), static_folder=os.path.join(BASE_DIR, 'static'))
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

PORT = 8085

def free_port(port):
    if sys.platform == 'win32':
        try:
            current_pid = os.getpid()
            out = subprocess.check_output('netstat -ano', shell=True).decode('utf-8', errors='ignore')
            for line in out.splitlines():
                if f':{port}' in line and 'LISTENING' in line:
                    parts = line.strip().split()
                    pid = int(parts[-1])
                    if pid != current_pid and pid != 0:
                        subprocess.run(f'taskkill /F /PID {pid}', shell=True, capture_output=True)
            time.sleep(0.5)
        except Exception:
            pass

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/managers', methods=['GET'])
def get_managers():
    mgrs = db_manager.get_managers_list()
    return jsonify({"status": "success", "managers": mgrs})

@app.route('/api/stats', methods=['GET'])
def get_stats():
    role = request.args.get('role', 'ADMIN').upper().strip()
    manager = request.args.get('manager', '').strip()
    stats = db_manager.get_stats(role, manager)
    return jsonify({"status": "success", "stats": stats})

@app.route('/api/transactions', methods=['GET'])
def get_transactions():
    role = request.args.get('role', 'ADMIN').upper().strip()
    manager = request.args.get('manager', '').strip()
    status_filter = request.args.get('status', '').strip() or None
    search_query = request.args.get('search', '').strip() or None
    bank_filter = request.args.get('bank', '').strip() or None
    amount_filter = request.args.get('amount', '').strip() or None
    date_filter = request.args.get('date', '').strip() or None
    
    txs = db_manager.get_transactions(role, manager, status_filter, search_query, bank_filter, amount_filter, date_filter)
    return jsonify({
        "status": "success",
        "total": len(txs),
        "transactions": txs
    })

@app.route('/api/smart_match', methods=['POST'])
def api_smart_match():
    data = request.json or {}
    role = data.get('role', 'ADMIN').upper().strip()
    manager = data.get('manager', '').strip()
    amount = data.get('amount')
    partial_name = data.get('name', '').strip()
    partial_utr = data.get('utr', '').strip()
    bank = data.get('bank', '').strip() or None
    
    ranked = db_manager.smart_match_finder(role, manager, amount, partial_name, partial_utr, bank)
    return jsonify({
        "status": "success",
        "total": len(ranked),
        "candidates": ranked
    })

@app.route('/api/claim', methods=['POST'])
def claim_transaction():
    data = request.json or {}
    tx_id = data.get('tx_id')
    manager = data.get('manager', '').strip()
    cust_id = data.get('customer_id', '').strip()
    cust_name = data.get('customer_name', '').strip()
    portfolio = data.get('portfolio', '').strip()
    caller = data.get('caller_name', '').strip()
    mode = data.get('mode_of_pay', '').strip()
    
    if not tx_id or not manager or not cust_id or not cust_name or not caller:
        return jsonify({"status": "error", "message": "Please fill all required fields (Cust ID, Name, Portfolio, Caller)."}), 400
        
    ok, msg = db_manager.claim_transaction(tx_id, manager, cust_id, cust_name, portfolio, caller, mode)
    if ok:
        return jsonify({"status": "success", "message": msg})
    return jsonify({"status": "error", "message": msg}), 400

@app.route('/api/unclaim', methods=['POST'])
def unclaim_transaction():
    data = request.json or {}
    tx_id = data.get('tx_id')
    manager = data.get('manager', '').strip()
    role = data.get('role', 'MANAGER').upper().strip()
    
    if not tx_id:
        return jsonify({"status": "error", "message": "Transaction ID required"}), 400
        
    ok, msg = db_manager.unclaim_transaction(tx_id, manager, role)
    if ok:
        return jsonify({"status": "success", "message": msg})
    return jsonify({"status": "error", "message": msg}), 400

@app.route('/api/admin/upload_statement', methods=['POST'])
def admin_upload_statement():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded"}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({"status": "error", "message": "Empty filename"}), 400
        
    save_path = os.path.join(UPLOAD_DIR, f"admin_statement_{int(time.time())}_{file.filename}")
    file.save(save_path)
    
    try:
        txs = core_engine.parse_account_transactions(save_path)
        batch_date = request.form.get('batch_date', '').strip() or None
        inserted, dup = db_manager.insert_transactions(txs, batch_date)
        return jsonify({
            "status": "success",
            "message": f"Successfully loaded bank statement! New transactions: {inserted}, Already existing: {dup}",
            "inserted": inserted,
            "duplicates": dup,
            "total_parsed": len(txs)
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/export_handsoff', methods=['GET'])
def export_handsoff():
    role = request.args.get('role', 'ADMIN').upper().strip()
    manager = request.args.get('manager', '').strip()
    handoff_date = request.args.get('handoff_date', '').strip()
    
    # Get allocated transactions based on role
    # Admin exports ALL allocated transactions across managers
    # Manager exports ONLY their team's allocated transactions
    txs = db_manager.get_transactions(role, manager, status_filter='ALLOCATED')
    if not txs:
        return jsonify({"status": "error", "message": "No allocated records found to generate Handsoff."}), 400
        
    # Convert DB rows to Handsoff result format
    config = core_engine.load_config()
    portfolio_map = config.get("portfolio_mapping", {})
    manager_agency_map = config.get("agency_manager_mapping", {})
    default_agency = config.get("default_agency", "INH")
    
    if not handoff_date:
        handoff_date = time.strftime("%d-%m-%Y")
        
    results = []
    for t in txs:
        port_norm = (t["portfolio"] or "").strip().upper()
        port_info = portfolio_map.get(port_norm, {})
        port_source = port_info.get("source", port_norm)
        prod_type = port_info.get("product", "PL" if "PL" in port_norm else ("CC" if "CC" in port_norm else "PL"))
        
        mgr_norm = (t["claimed_by_manager"] or "").strip().upper()
        agency = manager_agency_map.get(mgr_norm, default_agency)
        if (t["caller_name"] or "").strip().upper() == "KP" or mgr_norm == "KP":
            agency = "KP"
            
        cheque_no = t["cheque_extracted"] or core_engine.extract_cheque_no(t["narration"], t["cheque_raw"])
        
        results.append({
            "PORTFOLIO": t["portfolio"],
            "PORTFOLIO TYPE/SOURCE": port_source,
            "PRODUCT TYPE": prod_type,
            "Loan ID": t["customer_id"],
            "LOAN NO": t["customer_id"],
            "Customer Name": t["customer_name"],
            "Mode of Pay": t["mode_of_pay"] or "UPI",
            "Date of Pay": t["tx_date"],
            "Reference No": t["ref_no"],
            "Amount": float(t["amount"]),
            "CHEQUE.no": cheque_no,
            "TRANSACTION DESCRIPTION": t["narration"],
            "HAND OFF DATE": handoff_date,
            "AGENCY NAME": agency,
            "CALLER NAME": t["caller_name"],
            "AGENCY MANAGER": t["claimed_by_manager"],
            "BANK": t["bank_name"],
            "REMARKS": ""
        })
        
    out_filename = f"Handsoff_Export_{int(time.time())}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    core_engine.write_handsoff_excel(results, out_path)
    
    return jsonify({
        "status": "success",
        "count": len(results),
        "download_url": f"/api/download/{out_filename}"
    })

@app.route('/api/export_paid', methods=['GET'])
def export_paid():
    role = request.args.get('role', 'ADMIN').upper().strip()
    manager = request.args.get('manager', '').strip()
    
    txs = db_manager.get_transactions(role, manager, status_filter='ALLOCATED')
    if not txs:
        return jsonify({"status": "error", "message": "No allocated records found to export Paid File."}), 400
        
    paid_rows = []
    for t in txs:
        paid_rows.append({
            "customer_id": t["customer_id"],
            "customer_name": t["customer_name"],
            "portfolio": t["portfolio"],
            "date_of_pay": t["tx_date"],
            "amount": float(t["amount"]),
            "mode": t["mode_of_pay"] or "UPI",
            "narration": t["narration"],
            "bank_name": t["bank_name"],
            "officer_name": t["caller_name"],
            "manager": t["claimed_by_manager"]
        })
        
    out_filename = f"Paid_File_Export_{int(time.time())}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    core_engine.write_paid_excel(paid_rows, out_path)
    
    return jsonify({
        "status": "success",
        "count": len(paid_rows),
        "download_url": f"/api/download/{out_filename}"
    })

@app.route('/api/config', methods=['GET', 'POST'])
def api_config():
    if request.method == 'POST':
        data = request.json or {}
        core_engine.save_config(data)
        return jsonify({"status": "success", "config": data})
    cfg = core_engine.load_config()
    return jsonify(cfg)

@app.route('/api/download/<filename>')
def download_file(filename):
    target = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(target):
        return send_file(target, as_attachment=True, download_name=filename)
    return jsonify({"status": "error", "message": "File not found"}), 404

def open_browser():
    webbrowser.open(f"http://localhost:{PORT}")

if __name__ == '__main__':
    free_port(PORT)
    print(f"Starting Handsoff Automation Studio on http://localhost:{PORT} ...")
    Timer(1.2, open_browser).start()
    app.run(host='0.0.0.0', port=PORT, debug=False)
