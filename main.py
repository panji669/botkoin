from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
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
# 1. MANAJEMEN KONFIGURASI & EXCEL
# ==========================================
class EngineSettings(BaseModel):
    api_id: int
    api_hash: str
    nomor_hp: str
    bot_target: str
    otp_timeout: int = 50
    otp_interval: int = 3
    hunt_max: int = 10
    hunt_delay: float = 1.0
    jeda_setelah_order: float = 3.0
    jeda_antar_aksi: float = 0.5
    delay_aksi: float = 1.5
    delay_repeat: float = 3.0
    list_cookie: str = ""

telegram_client = None
phone_code_hash = None 
listener_aktif = False  
auto_claim_aktif = False
sedang_proses_siklus = False

def baca_konfigurasi():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    except Exception as e:
        print(f"\n❌ Error membaca konfigurasi: {e}")
        return {}

def ambil_dan_hapus_cookie():
    """Mengambil cookie baris teratas dan menghapusnya secara permanen dari file konfigurasi."""
    try:
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
    except Exception as e:
        print(f"\n❌ Error mengambil cookie: {e}")
        return None

def inisialisasi_excel_jika_belum_ada():
    """Membuat file Excel dengan format persis 3 kolom (Header di baris 1)."""
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
                cell.border = openpyxl.styles.Border(
                    left=openpyxl.styles.Side(style='thin'), 
                    right=openpyxl.styles.Side(style='thin'), 
                    top=openpyxl.styles.Side(style='thin'), 
                    bottom=openpyxl.styles.Side(style='thin')
                )
            
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 15
            wb.save(EXCEL_REPORT_FILE)
    except Exception as e:
        print(f"\n❌ Gagal menginisialisasi Excel: {e}")

inisialisasi_excel_jika_belum_ada()

def catat_sukses_claim_ke_excel(username, cookie_str):
    """Mencatat ke Excel sekaligus mencegah duplikasi data berdasarkan Raw Cookie yang sama."""
    try:
        inisialisasi_excel_jika_belum_ada()
        wb = openpyxl.load_workbook(EXCEL_REPORT_FILE)
        ws = wb.active
        
        # PENGAMAN DUPLIKASI: Cek apakah cookie ini sudah pernah tercatat sebelumnya di Kolom B
        for row in range(2, ws.max_row + 1):
            existing_cookie = ws.cell(row=row, column=2).value
            if existing_cookie and cookie_str and str(existing_cookie).strip() == str(cookie_str).strip():
                print(f"\n⚠️ [DUPLIKAT DIABAIKAN]: Cookie untuk akun '{username}' sudah ada di Excel.")
                return # Batalkan penyimpanan jika sudah ada!

        next_row = ws.max_row + 1
        border_thin = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'), 
            right=openpyxl.styles.Side(style='thin'), 
            top=openpyxl.styles.Side(style='thin'), 
            bottom=openpyxl.styles.Side(style='thin')
        )
        
        # Kolom A = Username, Kolom B = Raw Cookie, Kolom C = 3000
        ws.cell(row=next_row, column=1, value=username).border = border_thin
        ws.cell(row=next_row, column=2, value=cookie_str).border = border_thin
        ws.cell(row=next_row, column=3, value=3000).border = border_thin
        
        wb.save(EXCEL_REPORT_FILE)
        print(f"\n📊 [EXCEL LOG]: Berhasil mencatat Username '{username}' dengan koin 3,000!")
    except Exception as e:
        print(f"\n❌ Gagal mencatat ke Excel: {e}")

async def jalankan_siklus_misi(bot_target):
    global auto_claim_aktif, sedang_proses_siklus
    
    if not auto_claim_aktif or sedang_proses_siklus:
        return
        
    sedang_proses_siklus = True
    config = baca_konfigurasi()
    delay_repeat = config.get('delay_repeat', 3.0)
    
    print(f"\n⏳ Menunggu {delay_repeat} detik sebelum memulai siklus baru...")
    await asyncio.sleep(delay_repeat)
    
    if not auto_claim_aktif:
        sedang_proses_siklus = False
        return

    sisa_cookie = config.get('list_cookie', '').strip()
    if not sisa_cookie:
        print("\n❌ AUTO CLAIM BERHENTI: Seluruh daftar cookie di antrean sudah habis.")
        auto_claim_aktif = False
        sedang_proses_siklus = False
        return

    print(f"\n🚀 Mengirim perintah /mission ke {bot_target}...")
    try:
        await telegram_client.send_message(bot_target, '/mission')
    except Exception as e:
        print(f"\n❌ Gagal mengirim pesan /mission: {e}")
    
    await asyncio.sleep(2.0)
    sedang_proses_siklus = False


# ==========================================
# 2. LISTENER TELEGRAM & FLOW OTOMATIS
# ==========================================
def daftarkan_listener(client, bot_target):
    global listener_aktif
    if listener_aktif:
        return
        
    @client.on(events.NewMessage(chats=bot_target, incoming=True))
    async def proses_pesan_masuk(event):
        global auto_claim_aktif, sedang_proses_siklus
        config = baca_konfigurasi()
        delay_aksi = config.get('delay_aksi', 1.5)
        
        teks_masuk = event.raw_text
        print(f"\n📥 Pesan masuk: {teks_masuk}")
        
        await asyncio.sleep(1)
        try:
            markup = await event.get_reply_markup()
        except:
            markup = None

        buttons = event.message.buttons or (markup.rows if markup else None)

        if buttons:
            print("\n🔘 Tombol/Badge terdeteksi pada pesan!")
            ditemukan_misi = False
            ditemukan_login = False
            
            for row in buttons:
                row_buttons = row.buttons if hasattr(row, 'buttons') else row
                for button in row_buttons:
                    btn_text = getattr(button, 'text', '')
                    if "Install ShopeePay" in btn_text:
                        ditemukan_misi = True
                    elif "Raw Cookie" in btn_text:
                        ditemukan_login = True
            
            if ditemukan_misi:
                print(f"\n⚡ Menunggu {delay_aksi}s lalu mengeksekusi klik 'Install ShopeePay'...")
                await asyncio.sleep(delay_aksi)
                try:
                    await event.click(0, 0)
                    print("\n✅ Berhasil klik Install ShopeePay via koordinat (0,0)!")
                except Exception as err:
                    print(f"\n❌ Gagal klik Install ShopeePay: {err}")
                    
            elif ditemukan_login:
                print(f"\n⚡ Menunggu {delay_aksi}s lalu mengeksekusi klik 'Raw Cookie'...")
                await asyncio.sleep(delay_aksi)
                try:
                    await event.click(0, 1)
                    print("\n✅ Berhasil klik Raw Cookie via koordinat (0,1)!")
                except Exception as err:
                    print(f"\n❌ Gagal klik Raw Cookie: {err}")
                    
        elif "Masukkan raw cookie" in teks_masuk:
            print(f"\n⚡ Bot meminta cookie. Menunggu {delay_aksi}s...")
            await asyncio.sleep(delay_aksi) 
            
            config_current = baca_konfigurasi()
            cookies_str_preview = config_current.get('list_cookie', '').strip()
            cookie_akurat = cookies_str_preview.split('\n')[0].strip() if cookies_str_preview else ""
            
            cookie = ambil_dan_hapus_cookie()
            if cookie:
                await event.reply(cookie) 
                print("\n✅ Cookie terkirim & baris terpakai telah dihapus dari antrean server!")
                client._cookie_sedang_diproses = cookie_akurat
            else:
                print("\n❌ Gagal: Antrean cookie kosong!")
                auto_claim_aktif = False
                
        # Flow C: Hasil Akhir Sesi Akun
        else:
            if auto_claim_aktif and not sedang_proses_siklus:
                await asyncio.sleep(1.5)
                try:
                    pesan_terakhir = await client.get_messages(bot_target, limit=1)
                    if pesan_terakhir:
                        teks_gabungan = pesan_terakhir[0].raw_text.replace('\n', ' | ')
                    else:
                        teks_gabungan = teks_masuk.replace('\n', ' | ')
                except:
                    teks_gabungan = teks_masuk.replace('\n', ' | ')

                print(f"\n🎯 [HASIL AKHIR SESI AKUN]: {teks_gabungan}")
                print("--------------------------------------------------")

                try:
                    username = "Unknown"
                    if "✅ Login:" in teks_gabungan:
                        temp = teks_gabungan.split("✅ Login:")[1].strip()
                        username = temp.split(" ")[0].split("|")[0].strip()

                    cookie_terpakai = getattr(client, '_cookie_sedang_diproses', 'Cookie Kosong / Manual')
                    
                    # PENGAMAN DUPLIKASI: Cek apakah cookie atau username ini sudah baru saja dicatat
                    id_unik_sesi = f"{username}_{cookie_terpakai}"
                    if getattr(client, '_terakhir_dicatat', '') != id_unik_sesi and username != "Unknown":
                        client._terakhir_dicatat = id_unik_sesi
                        catat_sukses_claim_ke_excel(username, cookie_terpakai)
                    elif username == "Unknown":
                        catat_sukses_claim_ke_excel("Akun_Processed", cookie_terpakai)
                        
                except Exception as parse_err:
                    print(f"⚠️ Gagal parsing data untuk excel: {parse_err}")
                
                await jalankan_siklus_misi(bot_target)
                    
    listener_aktif = True
    print(f"\n🎧 Listener aktif memantau pesan masuk dari: {bot_target}")


# ==========================================
# 3. ENDPOINT API FASTAPI
# ==========================================
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
        return FileResponse(
            EXCEL_REPORT_FILE, 
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename='laporan_klaim_koin_enterprise.xlsx'
        )
    return {"status": "error", "message": "File laporan belum tersedia."}

@app.post("/api/simpan_pengaturan")
async def simpan_pengaturan(settings: EngineSettings):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(settings.dict(), f, indent=4)
        return {"status": "success", "message": "Konfigurasi & Antrean Cookie Berhasil Disimpan Otomatis!"}
    except Exception as e:
        return {"status": "error", "message": f"Gagal menyimpan konfigurasi: {str(e)}"}

@app.post("/api/hubungkan")
async def hubungkan_telegram():
    global telegram_client, phone_code_hash
    config = baca_konfigurasi()
    if not config or not config.get('api_id'):
        return {"status": "error", "message": "Simpan konfigurasi API ID & Hash terlebih dahulu!"}
        
    try:
        if telegram_client:
            try:
                if telegram_client.is_connected():
                    await telegram_client.disconnect()
            except:
                pass
            telegram_client = None

        telegram_client = TelegramClient('sesi_web', int(config['api_id']), str(config['api_hash']))
        await telegram_client.connect()
        
        if not await telegram_client.is_user_authorized():
            kirim_otp = await telegram_client.send_code_request(config['nomor_hp'])
            phone_code_hash = kirim_otp.phone_code_hash
            return {"status": "otp_required", "message": "Kode OTP berhasil dikirim ke aplikasi Telegram Anda."}

        daftarkan_listener(telegram_client, config['bot_target'])
        return {"status": "connected", "message": "Koneksi Telegram berhasil diinisialisasi!"}
    except Exception as e:
        print(f"\n❌ Error Koneksi Telegram: {str(e)}")
        return {"status": "error", "message": f"Koneksi gagal: {str(e)}"}

class OTPVerification(BaseModel):
    kode_otp: str

@app.post("/api/verifikasi_otp")
async def verifikasi_otp(data: OTPVerification):
    global telegram_client, phone_code_hash
    config = baca_konfigurasi()
    try:
        kode_otp = data.kode_otp.strip()
        if not telegram_client:
            return {"status": "error", "message": "Sesi client belum diinisialisasi. Silakan klik hubungkan ulang."}
            
        await telegram_client.sign_in(config['nomor_hp'], kode_otp, phone_code_hash=phone_code_hash)
        daftarkan_listener(telegram_client, config['bot_target'])
        return {"status": "success", "message": "Verifikasi OTP & Login Berhasil!"}
    except Exception as e:
        print(f"\n❌ Error Verifikasi OTP: {str(e)}")
        return {"status": "error", "message": f"Verifikasi gagal: {str(e)}"}

@app.post("/api/start_claim")
async def start_claim():
    global auto_claim_aktif, sedang_proses_siklus, telegram_client
    if not telegram_client or not await telegram_client.is_user_authorized():
        return {"status": "error", "message": "Belum terhubung ke Telegram! Silakan klik 'Hubungkan Sesi' terlebih dahulu."}
        
    config = baca_konfigurasi()
    sisa_cookie = config.get('list_cookie', '').strip()
    if not sisa_cookie:
        return {"status": "error", "message": "Antrean cookie masih kosong! Harap isi atau impor file .txt terlebih dahulu."}
        
    auto_claim_aktif = True
    sedang_proses_siklus = False
    
    print("\n▶️ AUTOMATION STARTED (Runner diaktifkan)")
    asyncio.create_task(jalankan_siklus_misi(config['bot_target']))
    return {"status": "success", "message": "Automation runner berhasil dijalankan."}

@app.post("/api/stop_claim")
async def stop_claim():
    global auto_claim_aktif, sedang_proses_siklus
    auto_claim_aktif = False
    sedang_proses_siklus = False
    print("\n⏹️ AUTOMATION STOPPED (Runner dihentikan)")
    return {"status": "success", "message": "Automation berhasil dihentikan."}