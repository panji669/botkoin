from fastapi import FastAPI, Request, HTTPException, Depends
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from telethon import TelegramClient, events
import asyncio
import json
import os
import sys
import io
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

CONFIG_FILE = 'settings.json'
USERS_FILE = 'users_db.json'
EXCEL_REPORT_FILE = 'laporan_klaim_koin_enterprise.xlsx'

# ==========================================
# 0. SISTEM BUFFER LIVE LOG TERMINAL
# ==========================================
class LogBuffer(io.StringIO):
    def __init__(self):
        super().__init__()
        self.log_content = []

    def write(self, s):
        sys.__stdout__.write(s)
        sys.__stdout__.flush()
        if s.strip():
            formatted_s = s if s.endswith('\n') else s + '\n'
            self.log_content.append(formatted_s)
            if len(self.log_content) > 150: 
                self.log_content.pop(0)

    def flush(self):
        pass

    def get_logs(self):
        return "".join(self.log_content)

log_stream = LogBuffer()
sys.stdout = log_stream


# ==========================================
# 1. MANAJEMEN DATABASE & EXCEL
# ==========================================
def baca_database_user():
    if not os.path.exists(USERS_FILE):
        default_db = {
            "owner": {
                "password": "ownerpassword",
                "role": "owner",
                "status": "aktif",
                "saldo": 100000.0
            }
        }
        with open(USERS_FILE, 'w') as f:
            json.dump(default_db, f, indent=4)
        return default_db
    try:
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    except:
        return {}

def simpan_database_user(db):
    with open(USERS_FILE, 'w') as f:
        json.dump(db, f, indent=4)

def baca_konfigurasi():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    except:
        return {}

def ambil_dan_hapus_cookie(username):
    try:
        db = baca_database_user()
        # Mengambil cookie dari konfigurasi user atau global config
        config = baca_konfigurasi()
        cookies_str = config.get('list_cookie', '').strip()
        if not cookies_str:
            return None
        cookies_list = cookies_str.split('\n')
        cookie_terpakai = cookies_list.pop(0)
        config['list_cookie'] = '\n'.join([c.strip() for c in cookies_list if c.strip()])
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
        return cookie_terpakai.strip()
    except:
        return None

def inisialisasi_excel_jika_belum_ada():
    try:
        if not os.path.exists(EXCEL_REPORT_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Log Klaim Sukses"
            headers = ["Username", "Raw Cookie", "Jumlah Koin"]
            header_fill = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            header_font = openpyxl.styles.Font(bold=True, size=11)
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 15
            wb.save(EXCEL_REPORT_FILE)
    except Exception as e:
        print(f"\n❌ Gagal membuat Excel: {e}")

inisialisasi_excel_jika_belum_ada()

def catat_sukses_claim_ke_excel(username, cookie_str):
    try:
        inisialisasi_excel_jika_belum_ada()
        wb = openpyxl.load_workbook(EXCEL_REPORT_FILE)
        ws = wb.active
        
        for row in range(2, ws.max_row + 1):
            existing_cookie = ws.cell(row=row, column=2).value
            if existing_cookie and cookie_str and str(existing_cookie).strip() == str(cookie_str).strip():
                return False

        next_row = ws.max_row + 1
        border_thin = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))
        
        ws.cell(row=next_row, column=1, value=username).border = border_thin
        ws.cell(row=next_row, column=2, value=cookie_str).border = border_thin
        ws.cell(row=next_row, column=3, value=3000).border = border_thin
        wb.save(EXCEL_REPORT_FILE)
        return True
    except Exception as e:
        print(f"\n❌ Gagal mencatat ke Excel: {e}")
        return False


# ==========================================
# 2. MULTI-USER TELEGRAM SESSIONS ENGINE
# ==========================================
active_clients = {}  # Format: {username: {"client": client, "phone_code_hash": hash, "auto_claim": False}}

async def jalankan_siklus_misi(username, bot_target):
    user_session = active_clients.get(username)
    if not user_session or not user_session.get("auto_claim"):
        return
        
    config = baca_konfigurasi()
    delay_repeat = config.get('delay_repeat', 3.0)
    await asyncio.sleep(delay_repeat)
    
    if not user_session.get("auto_claim"):
        return

    sisa_cookie = config.get('list_cookie', '').strip()
    if not sisa_cookie:
        print(f"\n❌ AUTO CLAIM [{username}] BERHENTI: Antrean cookie habis.")
        user_session["auto_claim"] = False
        return

    try:
        client = user_session["client"]
        await client.send_message(bot_target, '/mission')
    except Exception as e:
        print(f"\n❌ Gagal kirim /mission [{username}]: {e}")

def daftarkan_listener_user(username, client, bot_target):
    @client.on(events.NewMessage(chats=bot_target, incoming=True))
    async def proses_pesan_masuk(event):
        user_session = active_clients.get(username)
        if not user_session or not user_session.get("auto_claim"):
            return

        config = baca_konfigurasi()
        delay_aksi = config.get('delay_aksi', 1.5)
        teks_masuk = event.raw_text
        print(f"\n📥 [{username}] Pesan masuk: {teks_masuk}")
        
        await asyncio.sleep(1)
        try:
            markup = await event.get_reply_markup()
        except:
            markup = None

        buttons = event.message.buttons or (markup.rows if markup else None)

        if buttons:
            ditemukan_misi, ditemukan_login = False, False
            for row in buttons:
                row_buttons = row.buttons if hasattr(row, 'buttons') else row
                for button in row_buttons:
                    btn_text = getattr(button, 'text', '')
                    if "Install ShopeePay" in btn_text: ditemukan_misi = True
                    elif "Raw Cookie" in btn_text: ditemukan_login = True
            
            if ditemukan_misi:
                await asyncio.sleep(delay_aksi)
                try: await event.click(0, 0)
                except: pass
            elif ditemukan_login:
                await asyncio.sleep(delay_aksi)
                try: await event.click(0, 1)
                except: pass
                    
        elif "Masukkan raw cookie" in teks_masuk:
            await asyncio.sleep(delay_aksi) 
            config_current = baca_konfigurasi()
            cookies_str_preview = config_current.get('list_cookie', '').strip()
            cookie_akurat = cookies_str_preview.split('\n')[0].strip() if cookies_str_preview else ""
            
            cookie = ambil_dan_hapus_cookie(username)
            if cookie:
                await event.reply(cookie) 
                client._cookie_sedang_diproses = cookie_akurat
            else:
                user_session["auto_claim"] = False
                
        else:
            if user_session.get("auto_claim"):
                await asyncio.sleep(1.5)
                try:
                    pesan_terakhir = await client.get_messages(bot_target, limit=1)
                    teks_gabungan = pesan_terakhir[0].raw_text.replace('\n', ' | ') if pesan_terakhir else teks_masuk.replace('\n', ' | ')
                except:
                    teks_gabungan = teks_masuk.replace('\n', ' | ')

                print(f"\n🎯 [{username}] [HASIL AKHIR]: {teks_gabungan}")
                try:
                    akun_target = "Unknown"
                    if "Login:" in teks_gabungan:
                        temp = teks_gabungan.split("Login:")[1].strip()
                        akun_target = temp.split(" ")[0].split("|")[0].strip()

                    cookie_user = getattr(client, '_cookie_sedang_diproses', 'N/A')
                    berhasil_catat = catat_sukses_claim_ke_excel(akun_target if akun_target != "Unknown" else "Akun_Processed", cookie_user)
                    
                    if berhasil_catat:
                        db = baca_database_user()
                        if username in db:
                            db[username]["saldo"] = max(0.0, db[username].get("saldo", 0.0) - 200.0)
                            simpan_database_user(db)
                            print(f"💰 [{username}] Saldo terpotong Rp 200 (Sisa: Rp {db[username]['saldo']})")
                        
                except Exception as parse_err:
                    print(f"⚠️ Error parsing [{username}]: {parse_err}")
                
                await jalankan_siklus_misi(username, bot_target)


# ==========================================
# 3. ENDPOINTS API REST
# ==========================================
class UserAuth(BaseModel):
    username: str
    password: str

class UserRegister(BaseModel):
    username: str
    password: str

class AdminManage(BaseModel):
    target_username: str
    status: str 
    tambah_saldo: float 

class AdminDelete(BaseModel):
    target_username: str

class ConnectTelegram(BaseModel):
    username: str
    api_id: int
    api_hash: str
    nomor_hp: str
    bot_target: str

class VerifyOTP(BaseModel):
    username: str
    kode_otp: str

class EngineSettings(BaseModel):
    api_id: int
    api_hash: str
    nomor_hp: str
    bot_target: str
    delay_aksi: float = 1.5
    delay_repeat: float = 3.0
    list_cookie: str = ""

@app.get("/", response_class=HTMLResponse)
async def baca_index():
    if os.path.exists("index.html"):
        with open("index.html", "r", encoding="utf-8") as f:
            return f.read()
    return "<h3>File index.html belum di-upload.</h3>"

@app.post("/api/login")
async def login_user(data: UserAuth):
    db = baca_database_user()
    if data.username in db and db[data.username]["password"] == data.password:
        user_data = db[data.username]
        if user_data["role"] != "owner" and user_data.get("status") != "aktif":
            return {"status": "error", "message": "Akun Anda belum aktif! Silakan contact Admin untuk berlangganan."}
        return {"status": "success", "role": user_data["role"], "saldo": user_data["saldo"], "username": data.username}
    return {"status": "error", "message": "Username atau Password salah!"}

@app.post("/api/register")
async def register_user(data: UserRegister):
    db = baca_database_user()
    if data.username in db:
        return {"status": "error", "message": "Username sudah terdaftar!"}
    
    db[data.username] = {
        "password": data.password,
        "role": "user",
        "status": "pending",
        "saldo": 0.0
    }
    simpan_database_user(db)
    return {"status": "success", "message": "Registrasi berhasil! Akun Anda berstatus PENDING. Silakan contact Admin agar akun Anda diaktifkan."}

@app.get("/api/admin/users")
async def get_all_users(username: str):
    db = baca_database_user()
    if username not in db or db[username]["role"] != "owner":
        raise HTTPException(status_code=403, detail="Akses ditolak.")
    return {"users": db}

@app.post("/api/admin/manage")
async def admin_manage_user(data: AdminManage, admin_user: str):
    db = baca_database_user()
    if admin_user not in db or db[admin_user]["role"] != "owner":
        raise HTTPException(status_code=403, detail="Akses ditolak.")
    
    target = data.target_username
    if target not in db:
        return {"status": "error", "message": "User tidak ditemukan."}
    
    db[target]["status"] = data.status
    if data.tambah_saldo > 0:
        db[target]["saldo"] = db[target].get("saldo", 0.0) + data.tambah_saldo
        
    simpan_database_user(db)
    return {"status": "success", "message": f"Data user {target} berhasil diperbarui!"}

@app.post("/api/admin/delete")
async def admin_delete_user(data: AdminDelete, admin_user: str):
    db = baca_database_user()
    if admin_user not in db or db[admin_user]["role"] != "owner":
        raise HTTPException(status_code=403, detail="Akses ditolak.")
    
    target = data.target_username
    if target not in db:
        return {"status": "error", "message": "User tidak ditemukan."}
    if target == "owner":
        return {"status": "error", "message": "Akun owner utama tidak dapat dihapus!"}
    
    del db[target]
    simpan_database_user(db)
    return {"status": "success", "message": f"User {target} berhasil dihapus!"}

@app.get("/api/logs")
async def get_logs():
    return {"logs": log_stream.get_logs()}

@app.get("/api/get_config")
async def get_config():
    return baca_konfigurasi()

@app.get("/api/download_excel")
async def download_excel():
    inisialisasi_excel_jika_belum_ada()
    if os.path.exists(EXCEL_REPORT_FILE):
        return FileResponse(EXCEL_REPORT_FILE, filename='laporan_klaim_koin_enterprise.xlsx')
    return {"status": "error", "message": "File belum tersedia."}

@app.post("/api/simpan_pengaturan")
async def simpan_pengaturan(settings: EngineSettings):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(settings.dict(), f, indent=4)
    return {"status": "success", "message": "Konfigurasi disimpan!"}

@app.post("/api/hubungkan_user")
async def hubungkan_user_telegram(data: ConnectTelegram):
    try:
        session_name = f"sesi_{data.username}"
        if data.username in active_clients:
            try:
                if active_clients[data.username]["client"].is_connected():
                    await active_clients[data.username]["client"].disconnect()
            except:
                pass

        client = TelegramClient(session_name, int(data.api_id), str(data.api_hash))
        await client.connect()
        
        if not await client.is_user_authorized():
            res_otp = await client.send_code_request(data.nomor_hp)
            active_clients[data.username] = {
                "client": client,
                "phone_code_hash": res_otp.phone_code_hash,
                "auto_claim": False,
                "bot_target": data.bot_target
            }
            return {"status": "otp_required", "message": f"OTP terkirim ke Telegram [{data.username}]."}

        daftarkan_listener_user(data.username, client, data.bot_target)
        active_clients[data.username] = {
            "client": client,
            "phone_code_hash": None,
            "auto_claim": False,
            "bot_target": data.bot_target
        }
        return {"status": "connected", "message": f"Telegram [{data.username}] berhasil terhubung!"}
    except Exception as e:
        return {"status": "error", "message": f"Gagal koneksi: {str(e)}"}

@app.post("/api/verifikasi_otp_user")
async def verifikasi_otp_user(data: VerifyOTP):
    user_session = active_clients.get(data.username)
    if not user_session:
        return {"status": "error", "message": "Sesi tidak ditemukan. Hubungkan ulang."}
    
    config = baca_konfigurasi()
    try:
        client = user_session["client"]
        phone_hash = user_session["phone_code_hash"]
        
        # Ambil nomor HP dari settings atau config user
        await client.sign_in(phone=config.get('nomor_hp'), code=data.kode_otp.strip(), phone_code_hash=phone_hash)
        daftarkan_listener_user(data.username, client, user_session["bot_target"])
        user_session["phone_code_hash"] = None
        return {"status": "success", "message": f"Verifikasi OTP [{data.username}] Berhasil!"}
    except Exception as e:
        return {"status": "error", "message": f"Verifikasi gagal: {str(e)}"}

@app.post("/api/start_claim_user")
async def start_claim_user(data: dict):
    username = data.get("username")
    user_session = active_clients.get(username)
    if not user_session or not user_session["client"].is_connected():
        return {"status": "error", "message": "Telegram belum terhubung untuk user ini!"}
        
    db = baca_database_user()
    if username in db and db[username].get("saldo", 0) <= 0:
        return {"status": "error", "message": "Saldo Anda habis! Silakan lakukan pengisian saldo (top-up) ke Admin."}

    user_session["auto_claim"] = True
    asyncio.create_task(jalankan_siklus_misi(username, user_session["bot_target"]))
    return {"status": "success", "message": f"Automation [{username}] berhasil dijalankan."}

@app.post("/api/stop_claim_user")
async def stop_claim_user(data: dict):
    username = data.get("username")
    if username in active_clients:
        active_clients[username]["auto_claim"] = False
    return {"status": "success", "message": f"Automation [{username}] dihentikan."}
