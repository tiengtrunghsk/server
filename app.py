# -*- coding: utf-8 -*-
"""
Unified Flask Server - Tích hợp:
1. Trang chủ / Health check
2. Xử lý từ vựng tiếng Trung (Tự động chạy hoặc chạy qua endpoint /generate-vocab)
3. API Chấm điểm tiếng Trung (/check, /test, /warmup)
4. API Xóa tài khoản Firebase (/api/delete-user)
"""

import os
import json
import traceback
import threading
from flask import Flask, request, jsonify
from flask_cors import CORS
from fuzzywuzzy import fuzz
from pypinyin import pinyin, Style
import re
from openpyxl import load_workbook

# --- FIREBASE SETUP ---
import firebase_admin
from firebase_admin import credentials, auth, firestore

firebase_json_str = os.environ.get("FIREBASE_CREDENTIALS_JSON")
if firebase_json_str:
    cred_dict = json.loads(firebase_json_str)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred)
else:
    if os.path.exists("serviceAccountKey.json"):
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
    else:
        print("⚠️ Cảnh báo: Chưa cấu hình thông tin xác thực Firebase Admin SDK!")

app = Flask(__name__)
CORS(app)

# ==========================================
# PHẦN 1: TỪ VỰNG TIẾNG TRUNG (Excel to JSON)
# ==========================================
@app.route('/generate-vocab', methods=['GET', 'POST'])
def generate_vocab():
    try:
        EXCEL_PATH = "data/tu_vung_hsk.xlsx"
        OUTPUT_PATH = "data/vocab_meanings.json"

        if not os.path.exists(EXCEL_PATH):
            return jsonify({"success": False, "error": f"Không tìm thấy {EXCEL_PATH}"}), 404

        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        wb = load_workbook(EXCEL_PATH, data_only=True, read_only=True)
        result = {}

        for sheet_name in wb.sheetnames:
            if not sheet_name.upper().startswith("HSK"):
                continue
            ws = wb[sheet_name]
            is_hsk79 = any(x in sheet_name for x in ['7', '8', '9'])
            col_zh = 1
            col_vi = 4 if is_hsk79 else 5
            count_sheet = 0
            for row in ws.iter_rows(min_row=3, values_only=True):
                if not row or len(row) <= col_vi:
                    continue
                zh = str(row[col_zh]).strip() if row[col_zh] else ""
                vi = str(row[col_vi]).strip() if len(row) > col_vi and row[col_vi] else ""
                if zh and vi and zh not in result:
                    result[zh] = vi
                    count_sheet += 1

        wb.close()

        with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, separators=(",", ":"))

        size_kb = os.path.getsize(OUTPUT_PATH) / 1024
        return jsonify({"success": True, "total_words": len(result), "size_kb": round(size_kb, 1)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# ==========================================
# PHẦN 2: CHẤM ĐIỂM TIẾNG TRUNG (MacBert)
# ==========================================
_corrector = None
_corrector_lock = threading.Lock()
_corrector_failed = False

def get_corrector():
    global _corrector, _corrector_failed
    if _corrector is not None:
        return _corrector
    if _corrector_failed:
        return None

    with _corrector_lock:
        if _corrector is not None:
            return _corrector
        if _corrector_failed:
            return None

        try:
            print("[API] Loading pycorrector...")
            from pycorrector import MacBertCorrector
            _corrector = MacBertCorrector("shibing624/macbert4csc-base-chinese")
            print("[API] Loaded OK")
            return _corrector
        except Exception as e:
            print("[API] Load failed: " + str(e))
            _corrector_failed = True
            return None

def clean_text(text):
    if not text:
        return ''
    return re.sub(r'[。，！？、；：""''「」『』（）《》〈〉【】〔〕\s]+', '', str(text))

def get_pinyin(char):
    if not char:
        return ''
    try:
        res = pinyin(char, style=Style.TONE)
        if res and res[0]:
            return res[0][0]
    except Exception:
        pass
    return ''

def add_pinyin_to_errors(errors):
    for e in errors:
        if e.get('user'):
            e['user_pinyin'] = get_pinyin(e['user'])
        if e.get('correct'):
            e['correct_pinyin'] = get_pinyin(e['correct'])
    return errors

def find_diff(user, correct):
    u = list(clean_text(user))
    c = list(clean_text(correct))
    errors = []
    m, n = len(u), len(c)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m + 1): dp[i][0] = i
    for j in range(n + 1): dp[0][j] = j

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if u[i-1] == c[j-1]:
                dp[i][j] = dp[i-1][j-1]
            else:
                dp[i][j] = 1 + min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1])

    i, j = m, n
    while i > 0 or j > 0:
        if i > 0 and j > 0 and u[i-1] == c[j-1]:
            i -= 1; j -= 1
        elif i > 0 and j > 0 and dp[i][j] == dp[i-1][j-1] + 1:
            errors.append({'type': 'wrong', 'position': j, 'input_position': i - 1, 'user': u[i-1], 'correct': c[j-1]})
            i -= 1; j -= 1
        elif j > 0 and dp[i][j] == dp[i][j-1] + 1:
            errors.append({'type': 'missing', 'position': j, 'input_position': i, 'correct': c[j-1]})
            j -= 1
        elif i > 0 and dp[i][j] == dp[i-1][j] + 1:
            errors.append({'type': 'extra', 'position': j, 'input_position': i - 1, 'user': u[i-1]})
            i -= 1
        else:
            break
    errors.reverse()
    return errors, dp[m][n]

def grade(user_answer, correct_answer):
    u = clean_text(user_answer)
    c = clean_text(correct_answer)

    if not u:
        return {'score': 0, 'status': 'empty', 'message': 'Bạn chưa gõ gì cả', 'errors': [], 'suggestion': correct_answer, 'used_ai': False}

    if u == c:
        return {'score': 100, 'status': 'correct', 'message': 'Đúng hoàn toàn', 'errors': [], 'suggestion': None, 'used_ai': False}

    errors = []
    used_ai = False
    corrector = get_corrector()
    if corrector:
        try:
            corrected, err_info = corrector.correct(u)
            if err_info:
                used_ai = True
                for e in err_info:
                    pos = e.get('pos', 0)
                    errors.append({'type': 'wrong', 'position': pos, 'input_position': pos, 'user': e.get('orig', ''), 'correct': e.get('correct', '')})
        except Exception as e:
            print("[API] Correct error: " + str(e))

    if not errors:
        errors, _ = find_diff(u, c)

    max_len = max(len(u), len(c))
    distance = len(errors)
    ratio = max(0, min(1, 1 - (distance / max_len))) if max_len > 0 else 0

    if ratio >= 0.9: status, message = 'correct', 'Đúng (sai nhỏ)'
    elif ratio >= 0.6: status, message = 'partial', 'Gần đúng'
    else: status, message = 'wrong', 'Sai'

    add_pinyin_to_errors(errors)
    return {'score': round(ratio * 100), 'status': status, 'message': message, 'errors': errors, 'suggestion': correct_answer if u != c else None, 'used_ai': used_ai}

@app.route('/health')
def health():
    return jsonify({'ok': True, 'model_loaded': _corrector is not None, 'model_failed': _corrector_failed})

@app.route('/warmup')
def warmup():
    get_corrector()
    return jsonify({'ok': True, 'loaded': _corrector is not None})

@app.route('/check', methods=['POST'])
def check():
    try:
        data = request.get_json(force=True, silent=True) or {}
        u = data.get('user_answer', '')
        c = data.get('correct_answer', '')
        if not c: return jsonify({'error': 'Missing correct_answer'}), 400
        return jsonify(grade(u, c))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/test')
def test_get():
    u = request.args.get('u', '')
    c = request.args.get('c', '')
    if not c: return jsonify({'error': 'Thieu tham so c', 'usage': 'Them ?u=cau_user&c=cau_dung'})
    return jsonify(grade(u, c))


# ==========================================
# PHẦN 3: XÓA TÀI KHOẢN FIREBASE
# ==========================================
@app.route('/api/delete-user', methods=['POST'])
def delete_user():
    data = request.json or {}
    email = data.get('email')
    
    if not email:
        return jsonify({"success": False, "error": "Thiếu thông tin email"}), 400
        
    email_clean = email.strip().lower()
    
    try:
        db = firestore.client()
        deleted_auth = False
        deleted_firestore = False
        
        try:
            user_record = auth.get_user_by_email(email_clean)
            auth.delete_user(user_record.uid)
            deleted_auth = True
        except firebase_admin.auth.UserNotFoundError:
            pass
            
        doc_ref = db.collection('allowed_users').document(email_clean)
        if doc_ref.get().exists:
            doc_ref.delete()
            deleted_firestore = True
            
        if not deleted_auth and not deleted_firestore:
            return jsonify({"success": False, "error": f"Không tìm thấy user {email_clean} trong hệ thống."}), 404
            
        return jsonify({
            "success": True, 
            "message": f"Đã xóa hoàn toàn user {email_clean} (Auth: {deleted_auth}, Firestore: {deleted_firestore})"
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


# ==========================================
# HOME ENDPOINT
# ==========================================
@app.route('/')
def index():
    return jsonify({
        'status': 'running',
        'message': 'Unified Flask Server đang hoạt động!',
        'endpoints': {
            'health': '/health',
            'warmup': '/warmup',
            'check_api': 'POST /check',
            'test_api': 'GET /test?u=...&c=...',
            'generate_vocab': '/generate-vocab',
            'delete_user': 'POST /api/delete-user'
        }
    })

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
