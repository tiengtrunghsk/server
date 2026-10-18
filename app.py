# -*- coding: utf-8 -*-
"""
MEGA API 5-IN-1 (BẢN HOÀN CHỈNH - FIX SAFARI CORS)
Bao gồm: Chấm điểm, Từ vựng, OTP, Firebase Delete, Xuất Excel.
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
from pypinyin import pinyin, Style

import os
import sys
import json
import re
import random
import base64
import threading
import traceback
from datetime import datetime, timedelta

# Thư viện OTP & Firebase & Excel
from pymongo import MongoClient
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import firebase_admin
from firebase_admin import credentials, auth, firestore
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__)

# ==========================================================
# CẤU HÌNH CORS TỐI ĐA (DÀNH RIÊNG CHO SAFARI/IOS)
# ==========================================================
CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["GET", "POST", "OPTIONS", "PUT", "DELETE"],
        "allow_headers": ["Content-Type", "Authorization", "Accept", "Origin", "X-Requested-With", "X-CSRF-Token"]
    }
})

# Bắt buộc phải có hàm này để Safari chấp nhận request POST kèm JSON
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization,Accept,Origin,X-Requested-With')
    response.headers.add('Access-Control-Allow-Methods', 'GET,POST,OPTIONS,PUT,DELETE')
    return response

# ==========================================================
# PHẦN 1: KHỞI TẠO FIREBASE
# ==========================================================
FIREBASE_READY = False
try:
    firebase_json_str = os.environ.get("FIREBASE_CREDENTIALS_JSON")
    if firebase_json_str:
        cred = credentials.Certificate(json.loads(firebase_json_str))
        firebase_admin.initialize_app(cred)
        FIREBASE_READY = True
    elif os.path.exists("serviceAccountKey.json"):
        cred = credentials.Certificate("serviceAccountKey.json")
        firebase_admin.initialize_app(cred)
        FIREBASE_READY = True
except Exception as e:
    print("[API] Firebase init failed:", e)

# ==========================================================
# PHẦN 2: KHỞI TẠO MONGODB (OTP)
# ==========================================================
MONGO_URI = os.environ.get("MONGODB_URI")
mongo_client = None
otp_collection = None
last_sent_map = {}

if MONGO_URI:
    try:
        mongo_client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
        otp_collection = mongo_client["test"]["otpcodes"]
        otp_collection.create_index("createdAt", expireAfterSeconds=600)
        print("[API] MongoDB init OK")
    except Exception as e:
        print("[API] MongoDB init failed:", e)

# ==========================================================
# PHẦN 3: VOCAB SETUP
# ==========================================================
HAS_VOCAB = False
try:
    from vocab_data.radical_analyzer import _split_components, get_radical_for_word
    from vocab_data.mnemonic_generator import generate_mnemonic
    HAS_VOCAB = True
except Exception as e:
    print("[VocabAPI] vocab_data load failed:", e)

_AI_MNEMONICS = {}
_AI_MNEMONICS_INDEX = {}
_VOCAB_MEANINGS = {}
_SIMILAR_CHARS = {}

def _find_ai_mnemonics_file():
    here = os.path.dirname(os.path.abspath(__file__))
    for p in [os.path.join(here, "data", "ai_mnemonics.json"), "data/ai_mnemonics.json"]:
        if os.path.exists(p): return p
    return None

def _build_ai_index(data):
    idx = {}
    for k, v in data.items():
        parts = k.split("|")
        if len(parts) == 3 and parts[2].strip(): idx[parts[2].strip()] = v
    return idx

def _do_load_ai_mnemonics():
    global _AI_MNEMONICS, _AI_MNEMONICS_INDEX
    path = _find_ai_mnemonics_file()
    if not path: return False
    try:
        with open(path, "r", encoding="utf-8") as f: _AI_MNEMONICS = json.load(f)
        _AI_MNEMONICS_INDEX = _build_ai_index(_AI_MNEMONICS)
        return True
    except: return False

def get_ai_mnemonic(char):
    if not _AI_MNEMONICS_INDEX: _do_load_ai_mnemonics()
    return _AI_MNEMONICS_INDEX.get(str(char or "").strip(), "")

def _find_vocab_meanings_file():
    here = os.path.dirname(os.path.abspath(__file__))
    for p in [os.path.join(here, "data", "vocab_meanings.json"), "data/vocab_meanings.json"]:
        if os.path.exists(p): return p
    return None

def _do_load_vocab_meanings():
    global _VOCAB_MEANINGS
    path = _find_vocab_meanings_file()
    if not path: return False
    try:
        with open(path, "r", encoding="utf-8") as f: _VOCAB_MEANINGS = json.load(f)
        return True
    except: return False

def get_meaning(char): return _VOCAB_MEANINGS.get(char, "")

def _do_load_similar_chars():
    global _SIMILAR_CHARS
    here = os.path.dirname(os.path.abspath(__file__))
    folder = os.path.join(here, "vocab_data")
    if os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.startswith("similar_chars") and fname.endswith(".json"):
                try:
                    with open(os.path.join(folder, fname), "r", encoding="utf-8") as f:
                        _SIMILAR_CHARS.update(json.load(f))
                except: pass

def get_similar_chars(char):
    if not char: return [], ""
    entry = _SIMILAR_CHARS.get(char)
    if not entry: return [], ""
    if isinstance(entry, list): return [str(c).strip() for c in entry if c], ""
    if isinstance(entry, dict): return [str(c).strip() for c in entry.get('similar', []) if c], entry.get('diff', '')
    return [], ""

_do_load_ai_mnemonics()
_do_load_vocab_meanings()
_do_load_similar_chars()

def get_pinyin_vocab(char):
    if not char: return ''
    try:
        res = pinyin(char, style=Style.TONE)
        if res and res[0]: return res[0][0]
    except: pass
    return ''

def analyze_char(char):
    if not char or len(char) != 1 or not ('\u4e00' <= char <= '\u9fff') or not HAS_VOCAB:
        return None
        
    try:
        result = {'char': char, 'pinyin': get_pinyin_vocab(char)}
        
        # 1. Lấy nghĩa tiếng Việt
        try:
            vi = get_meaning(char)
            if vi: result['vi'] = vi
        except Exception as e:
            print(f"[Analyze] get_meaning error '{char}': {e}")
        
        # 2. Lấy bộ thủ
        try:
            main_rad = get_radical_for_word(char)
            if main_rad: 
                result['radical'] = {k: main_rad.get(k, '') for k in ['zh', 'pinyin', 'meaning', 'strokes', 'position']}
        except Exception as e:
            print(f"[Analyze] get_radical error '{char}': {e}")
        
        # 3. Phân tích các thành phần cấu tạo
        try:
            comps = _split_components(char)
            if comps: 
                result['components'] = [{k: c.get(k, '') for k in ['zh', 'pinyin', 'meaning', 'position']} for c in comps[:8]]
        except Exception as e:
            print(f"[Analyze] split_components error '{char}': {e}")

        # 4. Lấy mẹo nhớ (ưu tiên AI, sau đó đến tĩnh)
        try:
            mnemonic = get_ai_mnemonic(char) or generate_mnemonic(char)
            if mnemonic: result['mnemonic'] = mnemonic
        except Exception as e:
            print(f"[Analyze] mnemonic error '{char}': {e}")
        
        # 5. Lấy các chữ dễ nhầm
        try:
            sim_chars, diff_text = get_similar_chars(char)
            if sim_chars:
                result['similar'] = {
                    'list': [{'char': c, 'pinyin': get_pinyin_vocab(c), 'meaning': get_meaning(c)} for c in sim_chars if len(c)==1],
                    'diff': diff_text
                }
        except Exception as e:
            print(f"[Analyze] similar_chars error '{char}': {e}")
            
        return result
        
    except Exception as e:
        # Nếu có lỗi cực kỳ nghiêm trọng, vẫn trả về dữ liệu cơ bản nhất
        print(f"[Analyze] CRITICAL error '{char}': {e}")
        return {'char': char, 'pinyin': get_pinyin_vocab(char)}

# ==========================================================
# PHẦN 4: CHẤM ĐIỂM
# ==========================================================
_corrector = None
_corrector_lock = threading.Lock()
_corrector_failed = False

def get_corrector():
    global _corrector, _corrector_failed
    if _corrector is not None: return _corrector
    if _corrector_failed: return None
    with _corrector_lock:
        if _corrector is not None: return _corrector
        if _corrector_failed: return None
        try:
            from pycorrector import MacBertCorrector
            _corrector = MacBertCorrector("shibing624/macbert4csc-base-chinese")
            return _corrector
        except Exception:
            _corrector_failed = True
            return None

def clean(text):
    if not text: return ''
    return re.sub(r'[。，！？、；：""''「」『』（）《》〈〉【】〔〕\s]+', '', str(text))

def get_pinyin(char):
    if not char: return ''
    try:
        res = pinyin(char, style=Style.TONE)
        if res and res[0]: return res[0][0]
    except: pass
    return ''

def add_pinyin_to_errors(errors):
    for e in errors:
        if e.get('user'): e['user_pinyin'] = get_pinyin(e['user'])
        if e.get('correct'): e['correct_pinyin'] = get_pinyin(e['correct'])
    return errors

def find_diff(user, correct):
    u, c = list(clean(user)), list(clean(correct))
    errors = []
    m, n = len(u), len(c)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m + 1): dp[i][0] = i
    for j in range(n + 1): dp[0][j] = j
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if u[i-1] == c[j-1]: dp[i][j] = dp[i-1][j-1]
            else: dp[i][j] = 1 + min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1])
    i, j = m, n
    while i > 0 or j > 0:
        if i > 0 and j > 0 and u[i-1] == c[j-1]: i -= 1; j -= 1
        elif i > 0 and j > 0 and dp[i][j] == dp[i-1][j-1] + 1:
            errors.append({'type': 'wrong', 'position': j, 'input_position': i-1, 'user': u[i-1], 'correct': c[j-1]})
            i -= 1; j -= 1
        elif j > 0 and dp[i][j] == dp[i][j-1] + 1:
            errors.append({'type': 'missing', 'position': j, 'input_position': i, 'correct': c[j-1]})
            j -= 1
        elif i > 0 and dp[i][j] == dp[i-1][j] + 1:
            errors.append({'type': 'extra', 'position': j, 'input_position': i-1, 'user': u[i-1]})
            i -= 1
        else: break
    errors.reverse()
    return errors, dp[m][n]

def grade(user_answer, correct_answer):
    u, c = clean(user_answer), clean(correct_answer)
    if not u: return {'score': 0, 'status': 'empty', 'message': 'Bạn chưa gõ gì cả', 'errors': [], 'suggestion': correct_answer, 'used_ai': False}
    if u == c: return {'score': 100, 'status': 'correct', 'message': 'Đúng hoàn toàn', 'errors': [], 'suggestion': None, 'used_ai': False}
    
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
        except: pass
        
    if not errors: errors, _ = find_diff(u, c)
    
    max_len = max(len(u), len(c))
    ratio = max(0, min(1, 1 - (len(errors) / max_len))) if max_len > 0 else 0
    
    if ratio >= 0.9: status, message = 'correct', 'Đúng (sai nhỏ)'
    elif ratio >= 0.6: status, message = 'partial', 'Gần đúng'
    else: status, message = 'wrong', 'Sai'
    
    add_pinyin_to_errors(errors)
    return {'score': round(ratio * 100), 'status': status, 'message': message, 'errors': errors, 'suggestion': correct_answer if u != c else None, 'used_ai': used_ai}

# ==========================================================
# PHẦN 5: OTP GMAIL
# ==========================================================
def send_gmail_html(to, subject, html):
    cid, csecret = os.environ.get("GOOGLE_CLIENT_ID"), os.environ.get("GOOGLE_CLIENT_SECRET")
    rtoken, sender = os.environ.get("GOOGLE_REFRESH_TOKEN"), os.environ.get("GOOGLE_SENDER_EMAIL")
    if not all([cid, csecret, rtoken, sender]): raise Exception("Thiếu biến môi trường Google OAuth2")
    
    creds = Credentials(None, refresh_token=rtoken, client_id=cid, client_secret=csecret, token_uri="https://oauth2.googleapis.com/token")
    service = build('gmail', 'v1', credentials=creds)
    
    sub_enc = "=?UTF-8?B?" + base64.b64encode(subject.encode('utf-8')).decode('utf-8') + "?="
    msg = f'From: "Account Verification" <{sender}>\nTo: {to}\nSubject: {sub_enc}\nMIME-Version: 1.0\nContent-Type: text/html; charset=UTF-8\nContent-Transfer-Encoding: base64\n\n{base64.b64encode(html.encode("utf-8")).decode("utf-8")}'
    encoded = base64.urlsafe_b64encode(msg.encode("utf-8")).decode("utf-8")
    
    service.users().messages().send(userId='me', body={'raw': encoded}).execute()

def build_otp_email_html(otp, device_name, email):
    return f"""
    <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;background:#f8fafc;padding:20px;border-radius:12px">
        <div style="background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:24px;border-radius:12px 12px 0 0;text-align:center">
            <h1 style="color:#fff;margin:0;font-size:22px">🔐 Xác thực đăng nhập</h1>
        </div>
        <div style="background:#fff;padding:28px 24px;border-radius:0 0 12px 12px">
            <p style="color:#475569;font-size:14px;margin:0 0 16px">Thiết bị đang đăng nhập vào tài khoản <b>{email}</b>:</p>
            <div style="background:#f1f5f9;border-radius:10px;padding:14px 18px;margin-bottom:20px"><b>{device_name}</b></div>
            <div style="text-align:center;margin:24px 0">
                <div style="display:inline-block;background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:18px 36px;border-radius:12px">
                    <div style="color:#fff;font-size:36px;font-weight:900;letter-spacing:8px;">{otp}</div>
                </div>
            </div>
        </div>
    </div>
    """

# ==========================================================
# PHẦN 6: ROUTES
# ==========================================================

@app.route('/api/delete-user', methods=['POST', 'OPTIONS'])
def delete_user():
    if request.method == 'OPTIONS':
        return jsonify({}), 200
        
    if not FIREBASE_READY: 
        return jsonify({"success": False, "error": "Firebase chưa cấu hình"}), 503
        
    data = request.json or {}
    email = data.get('email')
    if not email: 
        return jsonify({"success": False, "error": "Thiếu email"}), 400
        
    email_clean = email.strip().lower()
    try:
        db = firestore.client()
        deleted_auth, deleted_fs = False, False
        try:
            user_record = auth.get_user_by_email(email_clean)
            auth.delete_user(user_record.uid)
            deleted_auth = True
        except firebase_admin.auth.UserNotFoundError: 
            pass
            
        doc_ref = db.collection('allowed_users').document(email_clean)
        if doc_ref.get().exists:
            doc_ref.delete()
            deleted_fs = True
            
        if not deleted_auth and not deleted_fs: 
            return jsonify({"success": False, "error": f"Không tìm thấy user {email_clean}"}), 404
            
        return jsonify({"success": True, "message": f"Đã xóa user {email_clean}"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/device/send-otp', methods=['POST', 'OPTIONS'])
def send_otp():
    if request.method == 'OPTIONS':
        return jsonify({}), 200
        
    if otp_collection is None: 
        return jsonify({'ok': False, 'error': 'OTP service chưa sẵn sàng'}), 503
        
    try:
        data = request.get_json(force=True, silent=True) or {}
        email, did, dname = data.get('email', ''), data.get('deviceId', 'unknown'), data.get('deviceName', 'Thiết bị')
        if not email or '@' not in email: 
            return jsonify({'ok': False, 'error': 'Email không hợp lệ'}), 400
        
        el = email.lower().strip()
        now = datetime.utcnow()
        
        if el in last_sent_map and (now - last_sent_map[el]).total_seconds() < 60:
            wait = int(60 - (now - last_sent_map[el]).total_seconds())
            return jsonify({'ok': False, 'error': f'Vui lòng đợi {wait} giây', 'waitSec': wait}), 429
            
        otp_collection.delete_many({'email': el})
        otp = str(random.randint(100000, 999999))
        otp_collection.insert_one({'email': el, 'otp': otp, 'deviceId': did, 'deviceName': dname, 'attempts': 0, 'createdAt': now})
        
        send_gmail_html(el, 'Your OTP Verification Code', build_otp_email_html(otp, dname, el))
        last_sent_map[el] = now
        return jsonify({'ok': True, 'message': 'Đã gửi mã OTP'})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/api/device/verify-otp', methods=['POST'])
def verify_otp():
    if otp_collection is None: return jsonify({'ok': False, 'error': 'OTP service chưa sẵn sàng'}), 503
    try:
        data = request.get_json(force=True, silent=True) or {}
        email, otp = data.get('email', ''), data.get('otp', '')
        if not email or not otp: return jsonify({'ok': False, 'error': 'Thiếu email hoặc OTP'}), 400
        
        el = email.lower().strip()
        record = otp_collection.find_one({'email': el}, sort=[('createdAt', -1)])
        
        if not record: return jsonify({'ok': False, 'error': 'OTP không tồn tại hoặc đã hết hạn'}), 400
        if datetime.utcnow() - record['createdAt'] > timedelta(minutes=10):
            otp_collection.delete_many({'email': el})
            return jsonify({'ok': False, 'error': 'OTP đã hết hạn'}), 400
        if record.get('attempts', 0) >= 5:
            otp_collection.delete_many({'email': el})
            return jsonify({'ok': False, 'error': 'Quá nhiều lần thử'}), 429
            
        if str(record['otp']) != str(otp).strip():
            otp_collection.update_one({'_id': record['_id']}, {'$inc': {'attempts': 1}})
            return jsonify({'ok': False, 'error': 'Mã OTP không đúng', 'attemptsLeft': 5 - (record.get('attempts', 0) + 1)}), 400
            
        otp_collection.delete_many({'email': el})
        return jsonify({'ok': True, 'message': 'Xác thực thành công'})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/check', methods=['POST', 'OPTIONS'])
def check():
    if request.method == 'OPTIONS':
        return jsonify({}), 200
    try:
        data = request.get_json(force=True, silent=True) or {}
        u, c = data.get('user_answer', ''), data.get('correct_answer', '')
        if not c: return jsonify({'error': 'Missing correct_answer'}), 400
        return jsonify(grade(u, c))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/test')
def test_get():
    u, c = request.args.get('u', ''), request.args.get('c', '')
    if not c: return jsonify({'error': 'Thieu tham so c', 'usage': 'Them ?u=...&c=...'})
    return jsonify(grade(u, c))

@app.route('/analyze', methods=['POST', 'OPTIONS'])
def analyze():
    if request.method == 'OPTIONS':
        return jsonify({}), 200
    try:
        data = request.get_json(force=True, silent=True) or {}
        char = data.get('char', '').strip()
        if not char or len(char) != 1: 
            return jsonify({'ok': False, 'error': 'Invalid char'}), 400
            
        res = analyze_char(char)
        
        if res:
            return jsonify({'ok': True, 'data': res})
        else:
            # Trả về mã 200 thay vì 500 để Frontend không báo lỗi kết nối
            return jsonify({'ok': False, 'error': 'Not found', 'char': char}), 200
            
    except Exception as e:
        print(f"[API /analyze] Lỗi: {e}")
        # Trả về mã 200 kèm thông báo lỗi
        return jsonify({'ok': False, 'error': str(e), 'char': char}), 200

@app.route('/vocab_test')
def vocab_test():
    ch = request.args.get('char', '').strip()
    res = analyze_char(ch)
    return jsonify({'ok': True, 'data': res}) if res else jsonify({'ok': False, 'error': 'Not found'}), 500

@app.route('/api/admin/export-excel', methods=['GET', 'POST'])
def export_excel():
    EXCEL_PATH = "data/tu_vung_hsk.xlsx"
    OUTPUT_PATH = "data/vocab_meanings.json"
    
    if not os.path.exists(EXCEL_PATH):
        return jsonify({"success": False, "error": f"Không tìm thấy {EXCEL_PATH}"}), 404
        
    try:
        wb = load_workbook(EXCEL_PATH, data_only=True, read_only=True)
        result = {}
        for sheet_name in wb.sheetnames:
            if not sheet_name.upper().startswith("HSK"): continue
            ws = wb[sheet_name]
            is_hsk79 = any(x in sheet_name for x in ['7', '8', '9'])
            col_zh, col_vi = 1, (4 if is_hsk79 else 5)
            
            for row in ws.iter_rows(min_row=3, values_only=True):
                if not row or len(row) <= col_vi: continue
                zh = str(row[col_zh]).strip() if row[col_zh] else ""
                vi = str(row[col_vi]).strip() if len(row) > col_vi and row[col_vi] else ""
                if zh and vi and zh not in result:
                    result[zh] = vi
        wb.close()
        
        with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, separators=(",", ":"))
            
        _do_load_vocab_meanings()
        return jsonify({"success": True, "message": f"Đã xuất {len(result)} từ vựng"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/health')
def health():
    return jsonify({
        'ok': True,
        'model_loaded': _corrector is not None,
        'firebase': FIREBASE_READY,
        'mongo': otp_collection is not None,
        'has_vocab': HAS_VOCAB,
        'ai_mnemonics_loaded': len(_AI_MNEMONICS),
        'vocab_meanings_loaded': len(_VOCAB_MEANINGS),
        'similar_chars_loaded': len(_SIMILAR_CHARS)
    })

@app.route('/warmup')
def warmup():
    get_corrector()
    return jsonify({'ok': True, 'loaded': _corrector is not None})

@app.route('/')
def index():
    return jsonify({'name': 'Mega API 5-in-1', 'version': '7.2-NoCORS', 'status': 'Running'})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
