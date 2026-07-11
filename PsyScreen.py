import tkinter as tk
from tkinter import ttk, scrolledtext, colorchooser, messagebox, filedialog
from screeninfo import get_monitors
import threading
from PIL import Image, ImageTk
from pystray import MenuItem as item, Icon as icon
import json
import os
import traceback
import time
import locale
import io
import base64
import requests
import pandas as pd
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except ImportError:
    messagebox.showerror("Eksik Kütüphane", "Bu özellik için Python 3.9+ veya 'tzdata' kütüphanesi gereklidir.\nLütfen Python sürümünüzü güncelleyin veya 'pip install tzdata' komutunu çalıştırın.")

# Türkçe tarih/saat formatı için
try:
    locale.setlocale(locale.LC_TIME, 'tr_TR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Turkish_Turkey.1254')
    except locale.Error:
        pass

def send_telegram_message(message_text, notification_type='mola'):
    if notification_type == 'mola':
        if not (mola_telegram_var and mola_telegram_var.get()): return
    elif notification_type == 'hasta':
        if not (hasta_telegram_var and hasta_telegram_var.get()): return
    else: return
    profil_listesi = load_data(PROFILER_DOSYASI, []); secilen_ad = selected_profile.get()
    secilen_profil_obj = next((p for p in profil_listesi if p.get('ad') == secilen_ad), None)
    if not secilen_profil_obj: print("Geçerli bir profil seçilmedi, telegram mesajı gönderilemiyor."); return
    token = secilen_profil_obj.get("telegram_token"); chat_id = secilen_profil_obj.get("telegram_chat_id")
    if not token or not chat_id: print(f"'{secilen_ad}' profili için Telegram bilgileri eksik, mesaj gönderilemedi."); return
    api_url = f"https://api.telegram.org/bot{token}/sendMessage"; params = {'chat_id': chat_id, 'text': message_text}
    try:
        threading.Thread(target=requests.get, args=(api_url,), kwargs={'params': params, 'timeout': 10}).start()
        print(f"Telegram'a gönderilen mesaj ({notification_type}): {message_text}")
    except Exception as e: print(f"Telegram mesajı gönderilirken hata oluştu: {e}")

def send_log_to_notion(token, db_id, properties_payload):
    if not token or not db_id: return
    url = "https://api.notion.com/v1/pages"
    headers = { "Authorization": f"Bearer {token}", "Content-Type": "application/json", "Notion-Version": "2022-06-28" }
    data = { "parent": {"database_id": db_id}, "properties": properties_payload }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=10)
        if response.status_code == 200: print(f"Notion'a başarıyla loglandı (DB ID: ...{db_id[-4:]})")
        else: print(f"Notion API Hatası: {response.status_code} - {response.text}")
    except Exception as e: print(f"Notion'a gönderilirken hata oluştu: {e}")

def log_mola_activity(profil_adi, olay_turu, mola_tipi, planlanan_sure=None, gerceklesen_sure=None):
    if (log_mola_var and log_mola_var.get()):
        DOSYA_ADI = "mola_kayitlari.xlsx"
        try:
            yeni_kayit = { "Tarih ve Saat": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")], "Doktor Profili": [profil_adi], "Olay Türü": [olay_turu], "Mola Tipi": [mola_tipi], "Planlanan Süre": [planlanan_sure or 'N/A'], "Gerçekleşen Süre": [gerceklesen_sure or 'N/A'] }
            yeni_kayit_df = pd.DataFrame(yeni_kayit)
            if not os.path.exists(DOSYA_ADI): yeni_kayit_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            else:
                mevcut_kayitlar_df = pd.read_excel(DOSYA_ADI, engine='openpyxl')
                guncel_df = pd.concat([mevcut_kayitlar_df, yeni_kayit_df], ignore_index=True)
                guncel_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            print(f"Mola aktivitesi başarıyla Excel'e loglandı: {olay_turu}")
        except Exception as e: messagebox.showwarning("Loglama Hatası", f"Mola aktivitesi Excel'e kaydedilemedi.\nDosya başka bir programda açık olabilir.\n\nHata: {e}")
    
    if log_notion_mola_var and log_notion_mola_var.get():
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == profil_adi), None)
        if secilen_profil_obj:
            token = secilen_profil_obj.get("notion_token")
            db_id = secilen_profil_obj.get("notion_mola_db_id")
            gerceklesen_saniye = None
            if gerceklesen_sure:
                try:
                    dakika, saniye = map(int, gerceklesen_sure.split(':'))
                    gerceklesen_saniye = (dakika * 60) + saniye
                except: gerceklesen_saniye = None
            payload = { "Olay Türü": {"title": [{"text": {"content": olay_turu}}]}, "Doktor Profili": {"select": {"name": profil_adi}}, "Mola Tipi": {"select": {"name": mola_tipi}}, "Planlanan Süre": {"rich_text": [{"text": {"content": planlanan_sure or 'N/A'}}]}, "Tarih ve Saat": {"date": {"start": datetime.now(ZoneInfo("Europe/Istanbul")).isoformat()}} }
            if gerceklesen_saniye is not None: payload["Gerçekleşen Süre"] = {"number": gerceklesen_saniye}
            threading.Thread(target=send_log_to_notion, args=(token, db_id, payload), daemon=True).start()

def log_hasta_cagirma(profil_adi, hasta_adi_soyadi, olay):
    if (log_cagirma_var and log_cagirma_var.get()):
        DOSYA_ADI = "hasta_cagirma_kayitlari.xlsx"
        try:
            yeni_kayit = { "Tarih ve Saat": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")], "Doktor Profili": [profil_adi], "Hasta Adı Soyadı": [hasta_adi_soyadi], "Olay": [olay] }
            yeni_kayit_df = pd.DataFrame(yeni_kayit)
            if not os.path.exists(DOSYA_ADI): yeni_kayit_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            else:
                mevcut_kayitlar_df = pd.read_excel(DOSYA_ADI, engine='openpyxl')
                guncel_df = pd.concat([mevcut_kayitlar_df, yeni_kayit_df], ignore_index=True)
                guncel_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            print(f"Hasta aktivitesi başarıyla Excel'e loglandı: {olay}")
        except Exception as e: messagebox.showwarning("Loglama Hatası", f"Hasta aktivitesi Excel'e kaydedilemedi.\nDosya başka bir programda açık olabilir.\n\nHata: {e}")
    
    if log_notion_cagirma_var and log_notion_cagirma_var.get():
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == profil_adi), None)
        if secilen_profil_obj:
            token = secilen_profil_obj.get("notion_token")
            db_id = secilen_profil_obj.get("notion_cagirma_db_id")
            payload = { "Olay": {"title": [{"text": {"content": olay}}]}, "Doktor Profili": {"select": {"name": profil_adi}}, "Hasta Adı Soyadı": {"rich_text": [{"text": {"content": hasta_adi_soyadi}}]}, "Tarih ve Saat": {"date": {"start": datetime.now(ZoneInfo("Europe/Istanbul")).isoformat()}} }
            threading.Thread(target=send_log_to_notion, args=(token, db_id, payload), daemon=True).start()

def log_muayene_suresi(profil_adi, olay, planlanan_sure=None, gerceklesen_sure=None):
    if (log_muayene_var and log_muayene_var.get()):
        DOSYA_ADI = "muayene_sureleri.xlsx"
        try:
            yeni_kayit = { "Tarih ve Saat": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")], "Doktor Profili": [profil_adi], "Olay": [olay], "Planlanan Süre": [planlanan_sure or 'N/A'], "Gerçekleşen Süre": [gerceklesen_sure or 'N/A'] }
            yeni_kayit_df = pd.DataFrame(yeni_kayit)
            if not os.path.exists(DOSYA_ADI): yeni_kayit_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            else:
                mevcut_kayitlar_df = pd.read_excel(DOSYA_ADI, engine='openpyxl')
                guncel_df = pd.concat([mevcut_kayitlar_df, yeni_kayit_df], ignore_index=True)
                guncel_df.to_excel(DOSYA_ADI, index=False, engine='openpyxl')
            print(f"Muayene aktivitesi başarıyla Excel'e loglandı: {olay}")
        except Exception as e: messagebox.showwarning("Loglama Hatası", f"Muayene aktivitesi Excel'e kaydedilemedi.\nDosya başka bir programda açık olabilir.\n\nHata: {e}")
    
    if log_notion_muayene_var and log_notion_muayene_var.get():
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == profil_adi), None)
        if secilen_profil_obj:
            token = secilen_profil_obj.get("notion_token")
            db_id = secilen_profil_obj.get("notion_muayene_db_id")
            planlanan_saniye, gerceklesen_saniye = None, None
            if planlanan_sure:
                try: dakika, saniye = map(int, planlanan_sure.split(':')); planlanan_saniye = (dakika * 60) + saniye
                except: pass
            if gerceklesen_sure:
                try: dakika, saniye = map(int, gerceklesen_sure.split(':')); gerceklesen_saniye = (dakika * 60) + saniye
                except: pass
            payload = { "Olay": {"title": [{"text": {"content": olay}}]}, "Doktor Profili": {"select": {"name": profil_adi}}, "Tarih ve Saat": {"date": {"start": datetime.now(ZoneInfo("Europe/Istanbul")).isoformat()}} }
            if planlanan_saniye is not None: payload["Planlanan Süre"] = {"number": planlanan_saniye}
            if gerceklesen_saniye is not None: payload["Gerçekleşen Süre"] = {"number": gerceklesen_saniye}
            threading.Thread(target=send_log_to_notion, args=(token, db_id, payload), daemon=True).start()

# --- Dosya ve Varsayılan Veri Ayarları ---
PROFILER_DOSYASI = "profiller.json"
MESAJLAR_DOSYASI = "mesajlar.json"
HASTANELER_DOSYASI = "hastaneler.json"
VARSAYILAN_PROFILER = [{"ad": "Uzm. Dr. Emrah SONGUR", "icerik": "Uzm. Dr. Emrah SONGUR", "telegram_token": "", "telegram_chat_id": "", "api_url": "", "api_bearer": "", "notion_token": "", "notion_mola_db_id": "", "notion_cagirma_db_id": "", "notion_muayene_db_id": "", "mola_mesaji_adi": "", "mola_sonrasi_mesaji_adi": "", "hasta_var_mesaji_adi": ""}]
VARSAYILAN_MESAJLAR = [{"ad": "Poliklinik Anonsu", "icerik": "Psikiyatri Polikliniği"}, {"ad": "Mola Anonsu", "icerik": "Doktor kısa bir mola vermiştir.\nLütfen bekleyiniz."}, {"ad": "Hasta İçeride", "icerik": "HASTA İÇERİDE\nLütfen bekleyiniz."}]
VARSAYILAN_HASTANELER = [{"ad": "Merkez Hastanesi", "logo": "hastaneler/default_logo.png"}]

# --- Global Değişkenler ---
root, message_window, patient_frame_on_screen, tray_icon, active_logo_ref, active_preview_logo_ref = (None,) * 6
text_area, return_time_entry, bg_color_swatch, fg_color_swatch, selected_screen, preview_square_frame, patient_entry, patient_combobox, refresh_patient_btn, selected_profile, profile_menu, selected_message, message_menu, selected_hospital, hospital_menu, screen_menu, warning_time_entry, close_screen_button, g_main_message_label = (None,) * 19
mola_telegram_var, hasta_telegram_var, mola_telegram_check, hasta_telegram_check = (None,) * 4
pip_window = None
pip_mola_button, pip_patient_button, pip_mola_entry = (None,) * 3
pip_patient_combobox, pip_refresh_patient_btn, pip_call_patient_btn, pip_reset_patient_btn, pip_screen_button, pip_main_window_button = (None,) * 6
pip_clock_label, pip_clock_job_id = (None,) * 2
selected_patient_var = None
mask_patient_name_var = None
is_on_stopwatch_break = False; stopwatch_seconds = 0; stopwatch_job_id = None
g_patient_inside_total_seconds = 0; pip_timer_seconds = 0
PREVIEW_SIZE = 300; DEFAULT_BG_COLOR, DEFAULT_FG_COLOR = '#FFFFFF', '#000000'; selected_bg_color, selected_fg_color = DEFAULT_BG_COLOR, DEFAULT_FG_COLOR
timer_label, local_timer_label, start_stop_button, reset_button, countdown_job_id, external_clock_job_id, end_break_button = (None,) * 7
countdown_seconds_remaining = None; is_countdown_paused = False; is_warning_sent = False 
is_on_break = False; is_on_overtime_break = False; overtime_seconds = 0; overtime_job_id = None 
patient_inside_button, patient_inside_timer_entry, is_patient_inside, patient_inside_job_id = (None,) * 4; is_patient_inside = False
g_patient_display_job, g_patient_name_label, g_patient_data_from_api, call_patient_btn, reset_patient_btn = (None,) * 5
log_mola_var, log_cagirma_var, log_muayene_var = (None,) * 3 
log_notion_mola_var, log_notion_cagirma_var, log_notion_muayene_var = (None,) * 3 
log_notion_mola_check, log_notion_cagirma_check, log_notion_muayene_check = (None,) * 3 

# --- Yüzen Buton (PiP) Fonksiyonları ---
_drag_data = {"x": 0, "y": 0}
def start_drag(event): _drag_data["x"] = event.x; _drag_data["y"] = event.y
def do_drag(event):
    if pip_window: dx = event.x - _drag_data["x"]; dy = event.y - _drag_data["y"]; x = pip_window.winfo_x() + dx; y = pip_window.winfo_y() + dy; pip_window.geometry(f"+{x}+{y}")

def toggle_external_screen_from_pip():
    if message_window and message_window.winfo_exists():
        close_secondary_screen(keep_timer=False)
    else:
        show_message()

def toggle_main_window_from_pip():
    if root and root.winfo_viewable():
        hide_window()
    else:
        show_window()

def toggle_pip_window():
    global pip_window
    if pip_window and pip_window.winfo_exists(): pip_window.destroy(); pip_window = None
    else: create_pip_window()

def update_pip_clock(clock_widget):
    global pip_clock_job_id
    if clock_widget and clock_widget.winfo_exists():
        clock_widget.config(text=time.strftime('%H:%M'))
        pip_clock_job_id = root.after(1000, update_pip_clock, clock_widget)
    elif 'pip_clock_job_id' in globals() and pip_clock_job_id:
        root.after_cancel(pip_clock_job_id)
        pip_clock_job_id = None

def create_pip_window():
    global pip_window, pip_mola_button, pip_patient_button, pip_mola_entry
    global pip_patient_combobox, pip_refresh_patient_btn, pip_call_patient_btn, pip_reset_patient_btn, pip_clock_label, pip_screen_button, pip_main_window_button
    
    if pip_window and pip_window.winfo_exists(): pip_window.lift(); return
    pip_window = tk.Toplevel(root)
    pip_window.geometry("880x35")
    pip_window.resizable(False, False)
    pip_window.overrideredirect(True)
    pip_window.wm_attributes("-topmost", 1)

    left_frame = tk.Frame(pip_window)
    left_frame.pack(side="left", fill="y")

    pip_clock_label = tk.Label(left_frame, text="", font=("Helvetica", 16, "bold"), fg="white", bg="#007bff", width=5)
    pip_clock_label.pack(side="left", fill="y")
    update_pip_clock(pip_clock_label)

    drag_handle = tk.Frame(left_frame, bg="#dc3545", width=10)
    drag_handle.pack(side="left", fill="y")

    pip_clock_label.bind("<ButtonPress-1>", start_drag)
    pip_clock_label.bind("<B1-Motion>", do_drag)
    drag_handle.bind("<ButtonPress-1>", start_drag)
    drag_handle.bind("<B1-Motion>", do_drag)
    
    content_frame = ttk.Frame(pip_window, padding=(5, 5, 5, 4)); content_frame.pack(side="left", fill="both", expand=True)
    style = ttk.Style(pip_window)
    style.configure("PiP.TButton", font=("Helvetica", 9, "bold"), foreground="white", padding=2)
    style.configure("Yellow.PiP.TButton", background="#ffc107", foreground="black"); style.map("Yellow.PiP.TButton", background=[('active', '#e0a800')])
    style.configure("Green.PiP.TButton", background="#28a745", foreground="white"); style.map("Green.PiP.TButton", background=[('active', '#218838')])
    style.configure("Red.PiP.TButton", background="#dc3545", foreground="white"); style.map("Red.PiP.TButton", background=[('active', '#c82333')])
    style.configure("Blue.TButton", background="#007bff", foreground="white"); style.map("Blue.TButton", background=[('active', '#0056b3')])
    
    mola_frame = ttk.Frame(content_frame); mola_frame.pack(side="left", fill="y", padx=(0, 2))
    pip_mola_entry = ttk.Entry(mola_frame, width=7, font=("Helvetica", 10)); pip_mola_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
    pip_mola_button = ttk.Button(mola_frame, text="Mola", style="Yellow.PiP.TButton", command=toggle_pip_break); pip_mola_button.pack(side="left", fill="x", expand=True)
    
    hasta_frame = ttk.Frame(content_frame); hasta_frame.pack(side="left", fill="y", padx=(2, 5))
    pip_patient_button = ttk.Button(hasta_frame, text="Hasta Yok", style="Green.PiP.TButton", command=toggle_patient_inside); pip_patient_button.pack(fill="both", expand=True)
    
    pip_patient_call_frame = ttk.Frame(content_frame); pip_patient_call_frame.pack(side="left", fill="x", expand=True)
    pip_patient_combobox = ttk.Combobox(pip_patient_call_frame, font=("Helvetica", 10), textvariable=selected_patient_var)
    if patient_combobox: pip_patient_combobox['values'] = patient_combobox['values']
    pip_patient_combobox.pack(side="left", fill="x", expand=True)
    
    pip_refresh_patient_btn = ttk.Button(pip_patient_call_frame, text="🔄", width=3, style="Blue.TButton", command=fetch_and_populate_patients); pip_refresh_patient_btn.pack(side="left", padx=2)
    pip_call_patient_btn = ttk.Button(pip_patient_call_frame, text="Çağır", style="Red.TButton", command=call_patient); pip_call_patient_btn.pack(side="left", padx=(0,2))
    pip_reset_patient_btn = ttk.Button(pip_patient_call_frame, text="Sıfırla", style="Red.TButton", command=reset_patient_call); pip_reset_patient_btn.pack(side="left")

    pip_main_window_button = ttk.Button(content_frame, text="Ana Panel", style="Blue.TButton", command=toggle_main_window_from_pip)
    pip_main_window_button.pack(side="right", fill="y", padx=(5, 0))

    pip_screen_button = ttk.Button(content_frame, text="Ekranı Aç", style="Green.PiP.TButton", command=toggle_external_screen_from_pip)
    pip_screen_button.pack(side="right", fill="y", padx=(5, 0))

    update_pip_button_states(); set_patient_inside_state(is_patient_inside)

def toggle_pip_break():
    global is_on_break, is_on_stopwatch_break, stopwatch_seconds, is_on_overtime_break
    if is_on_break or is_on_stopwatch_break or is_on_overtime_break: reset_countdown(); return
    mola_suresi_str = pip_mola_entry.get().strip()
    if mola_suresi_str:
        return_time_entry.delete(0, tk.END); return_time_entry.insert(0, mola_suresi_str)
        start_stop_countdown()
    else:
        is_on_stopwatch_break = True; stopwatch_seconds = 0
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
        if secilen_profil_obj:
            mola_mesaji_adi = secilen_profil_obj.get("mola_mesaji_adi")
            if mola_mesaji_adi:
                mola_mesaji_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == mola_mesaji_adi), None)
                if mola_mesaji_obj: set_message_text(mola_mesaji_obj.get('icerik', '')); show_message()
            doktor_adi = secilen_profil_obj.get('ad', 'Doktor')
            send_telegram_message(f"{doktor_adi} süresiz bir molaya çıktı (PiP).", notification_type='mola')
            log_mola_activity(doktor_adi, "Mola Başladı", "Kronometre")
        run_stopwatch_break(); update_all_button_states(); update_pip_button_states()

def run_stopwatch_break():
    global stopwatch_seconds, stopwatch_job_id
    if not is_on_stopwatch_break: return
    mins, secs = divmod(stopwatch_seconds, 60); time_str = '-{:02d}:{:02d}'.format(mins, secs)
    if local_timer_label and local_timer_label.winfo_exists(): local_timer_label.config(text=time_str)
    if pip_mola_button and pip_mola_button.winfo_exists(): pip_mola_button.config(text=time_str)
    stopwatch_seconds += 1; stopwatch_job_id = root.after(1000, run_stopwatch_break)

def patient_inside_tick_handler():
    global pip_timer_seconds, patient_inside_job_id, g_patient_inside_total_seconds
    pip_timer_seconds += 1; mins, secs = divmod(pip_timer_seconds, 60); time_str = f"{mins:02d}:{secs:02d}"
    if pip_patient_button and pip_patient_button.winfo_exists(): pip_patient_button.config(text=time_str)
    if g_patient_inside_total_seconds > 0:
        if pip_patient_button and pip_patient_button.winfo_exists(): pip_patient_button.config(style="Red.PiP.TButton")
        if pip_timer_seconds >= g_patient_inside_total_seconds and not (pip_window and pip_window.winfo_exists()): set_patient_inside_state(False); return
    patient_inside_job_id = root.after(1000, patient_inside_tick_handler)

def toggle_patient_inside(): set_patient_inside_state(not is_patient_inside)

def set_patient_inside_state(is_inside):
    global is_patient_inside, patient_inside_job_id, pip_timer_seconds, g_patient_inside_total_seconds
    if patient_inside_job_id: root.after_cancel(patient_inside_job_id); patient_inside_job_id = None
    is_patient_inside = is_inside
    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    doktor_adi = secilen_profil_obj.get('ad', 'Doktor') if secilen_profil_obj else 'Doktor'
    planlanan_sure = patient_inside_timer_entry.get().strip()
    if is_inside:
        pip_timer_seconds = 0; timer_str = planlanan_sure
        if timer_str:
            try:
                minutes, seconds = map(int, timer_str.split(':')); g_patient_inside_total_seconds = (minutes * 60) + seconds
            except (ValueError, AttributeError): g_patient_inside_total_seconds = 0
        else: g_patient_inside_total_seconds = 0
        log_muayene_suresi(doktor_adi, "Muayene Başladı", planlanan_sure=timer_str)
        if secilen_profil_obj:
            hasta_var_mesaji_adi = secilen_profil_obj.get("hasta_var_mesaji_adi")
            if hasta_var_mesaji_adi:
                hasta_var_mesaj_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == hasta_var_mesaji_adi), None)
                if hasta_var_mesaj_obj: set_message_text(hasta_var_mesaj_obj.get('icerik', '')); show_message()
        if patient_inside_button and patient_inside_button.winfo_exists(): patient_inside_button.config(text="Hasta Var", style="Red.TButton")
        if pip_patient_button and pip_patient_button.winfo_exists(): pip_patient_button.config(style="Red.PiP.TButton")
        patient_inside_tick_handler()
    else:
        gerceklesen_sure_str = time.strftime('%M:%S', time.gmtime(pip_timer_seconds))
        log_muayene_suresi(doktor_adi, "Muayene Bitti", planlanan_sure=planlanan_sure, gerceklesen_sure=gerceklesen_sure_str)
        g_patient_inside_total_seconds, pip_timer_seconds = 0, 0
        if secilen_profil_obj:
            mola_sonrasi_mesaji_adi = secilen_profil_obj.get("mola_sonrasi_mesaji_adi")
            if mola_sonrasi_mesaji_adi:
                mola_sonrasi_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == mola_sonrasi_mesaji_adi), None)
                if mola_sonrasi_obj:
                    set_message_text(mola_sonrasi_obj.get('icerik', ''))
                    if message_window and message_window.winfo_exists(): show_message()
        if patient_inside_button and patient_inside_button.winfo_exists(): patient_inside_button.config(text="Hasta Yok", style="Green.TButton")
        if pip_patient_button and pip_patient_button.winfo_exists(): pip_patient_button.config(text="Hasta Yok", style="Green.PiP.TButton")
    update_all_button_states(); update_pip_button_states()

def load_data(dosya_adi, varsayilan):
    if not os.path.exists(dosya_adi) or os.path.getsize(dosya_adi) < 5: save_data(dosya_adi, varsayilan); return varsayilan
    try:
        with open(dosya_adi, 'r', encoding='utf-8') as f: return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError): return varsayilan
def save_data(dosya_adi, veri):
    with open(dosya_adi, 'w', encoding='utf-8') as f: json.dump(veri, f, ensure_ascii=False, indent=4)
def open_data_manager(title, file_path, columns, refresh_callback, default_data, has_browse=False, dropdown_fields=None):
    if dropdown_fields is None: dropdown_fields = {}
    manager = tk.Toplevel(root); manager.title(title); manager.geometry("800x600"); manager.grab_set()
    current_items_list = load_data(file_path, default_data)
    tv_frame = ttk.Frame(manager, padding=10); tv_frame.pack(fill="both", expand=True)
    column_ids = list(columns.keys())
    tv = ttk.Treeview(tv_frame, columns=column_ids, show='headings')
    for col_id, col_text in columns.items(): tv.heading(col_id, text=col_text); tv.column(col_id, width=150)
    tv.pack(side="left", fill="both", expand=True)
    def populate_treeview():
        for i in tv.get_children(): tv.delete(i)
        for item in current_items_list: values = [str(item.get(col_id, '')).replace("\n", " ↵ ") for col_id in column_ids]; tv.insert('', 'end', values=values)
    populate_treeview()
    entry_frame = ttk.Frame(manager, padding=10); entry_frame.pack(fill="x", expand=True)
    entry_widgets = {}; string_vars = {}
    for i, (col_id, col_text) in enumerate(columns.items()):
        ttk.Label(entry_frame, text=f"{col_text}:").grid(row=i, column=0, sticky="w", padx=5, pady=2)
        if col_id in dropdown_fields:
            mesaj_listesi_guncel = load_data(MESAJLAR_DOSYASI, []); options = [m.get('ad') for m in mesaj_listesi_guncel]; var = tk.StringVar(manager)
            entry_widget = ttk.OptionMenu(entry_frame, var, "", *options); string_vars[col_id] = var
        elif 'içerik' in col_text.lower(): entry_widget = scrolledtext.ScrolledText(entry_frame, height=4, font=("Helvetica", 9))
        else: entry_widget = ttk.Entry(entry_frame)
        entry_widget.grid(row=i, column=1, sticky="ew", padx=5); entry_widgets[col_id] = entry_widget
        if col_id == 'logo' and has_browse:
            def browse_factory(e=entry_widget):
                filepath = filedialog.askopenfilename(initialdir="hastaneler", title="Logo Seç", filetypes=(("Resim Dosyaları", "*.png;*.jpg;*.jpeg"),("Tüm Dosyalar", "*.*")))
                if filepath: e.delete(0, tk.END); e.insert(0, os.path.relpath(filepath).replace("\\", "/"))
            browse_btn = ttk.Button(entry_frame, text="Gözat...", command=browse_factory); browse_btn.grid(row=i, column=2, padx=5)
    entry_frame.columnconfigure(1, weight=1)
    def on_select(e):
        if not tv.focus(): return
        selected_iid = tv.focus(); index = tv.index(selected_iid); item_data = current_items_list[index]
        for col_id, widget in entry_widgets.items():
            content = item_data.get(col_id, "")
            if col_id in string_vars: string_vars[col_id].set(content)
            elif isinstance(widget, scrolledtext.ScrolledText): widget.delete("1.0", tk.END); widget.insert("1.0", content)
            else: widget.delete(0, tk.END); widget.insert(0, content)
    tv.bind('<<TreeviewSelect>>', on_select)
    btn_frame = ttk.Frame(manager, padding=10); btn_frame.pack(fill="x")
    def add():
        new_item = {col_id: (string_vars[col_id].get() if col_id in string_vars else (widget.get("1.0", tk.END).strip() if isinstance(widget, scrolledtext.ScrolledText) else widget.get().strip())) for col_id, widget in entry_widgets.items()}
        if any(col in ['ad', 'icerik'] and not new_item[col] for col in new_item): messagebox.showwarning("Eksik Bilgi", "'Ad' ve 'İçerik' alanları boş olamaz.", parent=manager); return
        current_items_list.append(new_item); populate_treeview()
        for col_id, widget in entry_widgets.items():
            if col_id in string_vars: string_vars[col_id].set("")
            elif isinstance(widget, scrolledtext.ScrolledText): widget.delete("1.0", tk.END)
            else: widget.delete(0, tk.END)
    def update():
        if not tv.focus(): return
        index = tv.index(tv.focus()); updated_item = {col_id: (string_vars[col_id].get() if col_id in string_vars else (widget.get("1.0", tk.END).strip() if isinstance(widget, scrolledtext.ScrolledText) else widget.get().strip())) for col_id, widget in entry_widgets.items()}
        if any(col in ['ad', 'icerik'] and not updated_item[col] for col in updated_item): messagebox.showwarning("Eksik Bilgi", "'Ad' ve 'İçerik' alanları boş olamaz.", parent=manager); return
        current_items_list[index] = updated_item; populate_treeview()
    def delete():
        if not tv.selection(): return
        for i in sorted([tv.index(i) for i in tv.selection()], reverse=True): del current_items_list[i]
        populate_treeview()
    ttk.Button(btn_frame, text="Yeni Ekle", command=add).pack(side="left", expand=True)
    ttk.Button(btn_frame, text="Seçileni Güncelle", command=update).pack(side="left", expand=True)
    ttk.Button(btn_frame, text="Seçileni Sil", command=delete).pack(side="left", expand=True)
    def save(): save_data(file_path, current_items_list); refresh_callback(); manager.destroy()
    ttk.Button(manager, text="Kaydet ve Kapat", command=save, style="Blue.TButton").pack(pady=10)

def refresh_dropdown(menu_widget, string_var, data_list, key):
    menu = menu_widget["menu"]; menu.delete(0, "end")
    if not data_list: string_var.set("Liste Boş"); menu.add_command(label="Liste Boş", state="disabled"); return
    item_names = [item.get(key, '') for item in data_list]
    for name in item_names: menu.add_command(label=name, command=lambda v=name: string_var.set(v))
    if item_names: string_var.set(item_names[0])
def refresh_hospital_dropdown(): refresh_dropdown(hospital_menu, selected_hospital, load_data(HASTANELER_DOSYASI, []), 'ad')
def refresh_profile_dropdown(): refresh_dropdown(profile_menu, selected_profile, load_data(PROFILER_DOSYASI, []), 'ad')
def refresh_message_dropdown(): refresh_dropdown(message_menu, selected_message, load_data(MESAJLAR_DOSYASI, []), 'ad')
def profile_changed(*args):
    update_preview()
    if not all([mola_telegram_check, hasta_telegram_check, log_notion_mola_check, log_notion_cagirma_check, log_notion_muayene_check]): return
    
    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    
    # Telegram Ayarları
    telegram_state = 'disabled'
    if secilen_profil_obj:
        token, chat_id = secilen_profil_obj.get("telegram_token"), secilen_profil_obj.get("telegram_chat_id")
        if token and chat_id: telegram_state = "normal"
    
    mola_telegram_check.config(state=telegram_state)
    hasta_telegram_check.config(state=telegram_state)
    if telegram_state == 'disabled':
        mola_telegram_var.set(False)
        hasta_telegram_var.set(False)

    # Notion Ayarları
    notion_token_exists = secilen_profil_obj and secilen_profil_obj.get("notion_token")
    
    # Mola Notion Checkbox
    mola_db_exists = notion_token_exists and secilen_profil_obj.get("notion_mola_db_id")
    log_notion_mola_check.config(state='normal' if mola_db_exists else 'disabled')
    if not mola_db_exists: log_notion_mola_var.set(False)
    
    # Çağırma Notion Checkbox
    cagirma_db_exists = notion_token_exists and secilen_profil_obj.get("notion_cagirma_db_id")
    log_notion_cagirma_check.config(state='normal' if cagirma_db_exists else 'disabled')
    if not cagirma_db_exists: log_notion_cagirma_var.set(False)

    # Muayene Notion Checkbox
    muayene_db_exists = notion_token_exists and secilen_profil_obj.get("notion_muayene_db_id")
    log_notion_muayene_check.config(state='normal' if muayene_db_exists else 'disabled')
    if not muayene_db_exists: log_notion_muayene_var.set(False)
    
    # API Ayarları
    if secilen_profil_obj and secilen_profil_obj.get("api_url"):
        patient_entry.grid_remove()
        patient_combobox.grid(row=0, column=0, sticky='ew')
        refresh_patient_btn.grid(row=0, column=1)
        fetch_and_populate_patients()
    else:
        patient_combobox.grid_remove()
        refresh_patient_btn.grid_remove()
        patient_entry.grid(row=0, column=0, sticky='ew')

def message_changed(*args):
    secilen_mesaj_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == selected_message.get()), None)
    if secilen_mesaj_obj: set_message_text(secilen_mesaj_obj.get('icerik', ''))
def set_message_text(content):
    if text_area: text_area.config(fg="black"); text_area.delete("1.0", tk.END); text_area.insert("1.0", content); update_preview()
def turkce_upper(text): return text.replace('i', 'İ').upper()
def censor_name(full_name):
    if not full_name.strip(): return ""
    return " ".join([p[:2] + '*' * (len(p) - 2) if len(p) > 2 else p for p in turkce_upper(full_name).split()])
def fetch_and_populate_patients():
    global patient_combobox, g_patient_data_from_api, pip_patient_combobox
    profil_listesi = load_data(PROFILER_DOSYASI, []); secilen_profil_obj = next((p for p in profil_listesi if p.get('ad') == selected_profile.get()), None)
    api_url = secilen_profil_obj.get("api_url") if secilen_profil_obj else None
    def set_combobox_values(values, default_text):
        if patient_combobox: patient_combobox['values'] = values
        if pip_window and pip_window.winfo_exists() and pip_patient_combobox: pip_patient_combobox['values'] = values
        selected_patient_var.set(default_text)
    if not api_url: set_combobox_values([], "API URL'si ayarlanmamış."); return
    try:
        selected_patient_var.set("Hasta listesi yükleniyor..."); root.update_idletasks()
        bearer = (secilen_profil_obj.get("api_bearer") or "").strip()
        headers = {}
        if bearer:
            headers["Authorization"] = f"Bearer {bearer}"
        response = requests.get(api_url, timeout=10, headers=headers)
        if response.status_code == 200:
            g_patient_data_from_api = response.json()
            if g_patient_data_from_api:
                display_list = [item['display_text'] for item in g_patient_data_from_api]
                set_combobox_values(display_list, display_list[0] if display_list else "")
            else: g_patient_data_from_api = []; set_combobox_values([], "Bugün için randevu bulunamadı.")
        else:
            g_patient_data_from_api = []
            set_combobox_values([], "Liste yüklenemedi. API Hatası.")
            messagebox.showerror("API Hatası", f"Sunucudan hata kodu alındı: {response.status_code}\n{response.text}", parent=root)
    except requests.exceptions.RequestException as e:
        g_patient_data_from_api = []
        set_combobox_values([], "Liste yüklenemedi. Bağlantı Hatası.")
        messagebox.showerror("Bağlantı Hatası", f"Hasta listesi alınamadı.\n\nHata: {e}", parent=root)
def patient_display_handler(seconds_left):
    global g_patient_display_job, g_patient_name_label
    if seconds_left <= 0 or not (g_patient_name_label and g_patient_name_label.winfo_exists()):
        reset_patient_call(clear_entry=False); return
    if seconds_left <= 5:
        current_color = g_patient_name_label.cget("fg")
        g_patient_name_label.config(fg="#dc3545" if current_color == "white" else "white")
    g_patient_display_job = root.after(500, lambda: patient_display_handler(seconds_left - 0.5))

def call_patient():
    global g_patient_name_label, g_patient_display_job, g_patient_data_from_api, mask_patient_name_var
    reset_patient_call(clear_entry=False)
    selected_text = selected_patient_var.get().strip()
    if not selected_text or "yükleniyor" in selected_text or "bulunamadı" in selected_text or "yüklenemedi" in selected_text or "ayarlanmamış" in selected_text:
        messagebox.showinfo("Bilgi", "Lütfen geçerli bir hasta seçin."); return
    patient_name_to_call = selected_text
    for item in g_patient_data_from_api:
        if item['display_text'] == selected_text: patient_name_to_call = item['name_only']; break
    else: patient_name_to_call = selected_text.split('(')[0].strip()
    doktor_adi = selected_profile.get()
    log_hasta_cagirma(doktor_adi, patient_name_to_call, "Hasta Çağrıldı")
    
    if mask_patient_name_var and mask_patient_name_var.get():
        display_name = censor_name(patient_name_to_call)
    else:
        display_name = turkce_upper(patient_name_to_call)
        
    send_telegram_message(f"Sn. {doktor_adi}, {display_name} isimli hastayı çağırdınız.", notification_type='hasta')
    container = tk.Frame(patient_frame_on_screen, bg=selected_bg_color); container.pack(pady=20)
    red_box = tk.Frame(container, bg="#dc3545"); red_box.pack()
    tk.Label(red_box, text="ÇAĞIRILAN HASTA", font=("Helvetica", 27, "bold"), fg="white", bg="#dc3545").pack(pady=(5,0), padx=20)
    g_patient_name_label = tk.Label(red_box, text=display_name, font=("Helvetica", 60, "bold"), fg="white", bg="#dc3545"); g_patient_name_label.pack(pady=(5,10), padx=20)
    patient_display_handler(20); update_preview()

def reset_patient_call(clear_entry=True):
    global g_patient_display_job, g_patient_name_label
    if clear_entry and selected_patient_var and selected_patient_var.get():
        selected_text = selected_patient_var.get().strip()
        if selected_text and not ("yükleniyor" in selected_text or "bulunamadı" in selected_text):
            patient_name_to_log = selected_text
            for item in g_patient_data_from_api:
                if item['display_text'] == selected_text: patient_name_to_log = item['name_only']; break
            else: patient_name_to_log = selected_text.split('(')[0].strip()
            log_hasta_cagirma(selected_profile.get(), patient_name_to_log, "Ekrandan Temizlendi")
    if g_patient_display_job: root.after_cancel(g_patient_display_job); g_patient_display_job = None
    if patient_frame_on_screen:
        for widget in patient_frame_on_screen.winfo_children(): widget.destroy()
    g_patient_name_label = None
    if clear_entry and selected_patient_var: selected_patient_var.set('')
    update_preview()
def choose_color(target_widget, color_type):
    global selected_bg_color, selected_fg_color
    initial_color = selected_bg_color if color_type == 'bg' else selected_fg_color
    color_code = colorchooser.askcolor(title="Renk Seçin", color=initial_color)
    if color_code and color_code[1]:
        if color_type == 'bg': selected_bg_color = color_code[1]
        elif color_type == 'fg': selected_fg_color = color_code[1]
        update_preview()
def clear_form():
    global selected_bg_color, selected_fg_color
    if text_area: text_area.delete("1.0", tk.END)
    if return_time_entry: return_time_entry.delete(0, tk.END)
    if selected_patient_var: selected_patient_var.set('')
    if warning_time_entry: warning_time_entry.delete(0, tk.END)
    if patient_inside_timer_entry: patient_inside_timer_entry.delete(0, tk.END)
    selected_bg_color, selected_fg_color = DEFAULT_BG_COLOR, DEFAULT_FG_COLOR
    reset_countdown();
    if is_patient_inside: set_patient_inside_state(False)
    if root: update_preview()
def on_text_area_focus_in(event):
    if text_area and text_area.get("1.0", tk.END).strip().lower() == "notunuzu buraya yazın...": text_area.delete("1.0", tk.END); text_area.config(fg="black")
def on_text_area_focus_out(event):
    if text_area and not text_area.get("1.0", tk.END).strip(): text_area.config(fg="grey"); text_area.insert("1.0", "Notunuzu buraya yazın...")
def show_message():
    global message_window, patient_frame_on_screen, active_logo_ref, timer_label, external_clock_job_id, screen_menu, g_main_message_label, close_screen_button
    close_secondary_screen(keep_timer=True)
    selected_screen_name = selected_screen.get(); monitors = get_monitors(); chosen_monitor = next((m for i, m in enumerate(monitors) if selected_screen_name.startswith(f"Ekran {i+1}")), None)
    if not chosen_monitor: messagebox.showerror("Hata", "Geçerli bir ekran seçilemedi!"); return
    if screen_menu: screen_menu.config(style="White.TMenubutton")
    if close_screen_button: close_screen_button.config(text="Dış Ekran Açık", style="Red.TButton")
    message_window = tk.Toplevel(root); message_window.overrideredirect(True); message_window.geometry(f"{chosen_monitor.width}x{chosen_monitor.height}+{chosen_monitor.x}+{chosen_monitor.y}"); message_window.configure(bg=selected_bg_color)
    logo_frame = tk.Frame(message_window, bg=selected_bg_color); logo_frame.place(x=30, y=30, anchor="nw")
    hospital_name_label_on_screen = tk.Label(message_window, text="", font=("Helvetica", 48, "bold"), fg=selected_fg_color, bg=selected_bg_color); hospital_name_label_on_screen.place(x=chosen_monitor.width - 30, y=30, anchor="ne")
    main_content_frame = tk.Frame(message_window, bg=selected_bg_color); main_content_frame.place(relx=0.5, y=300, anchor="n")
    patient_frame_on_screen = tk.Frame(message_window, bg=selected_bg_color); patient_frame_on_screen.place(relx=0.5, y=chosen_monitor.height - 30, anchor="s")
    timer_label = tk.Label(message_window, text="", font=("Helvetica", 48, "bold"), fg=selected_fg_color, bg=selected_bg_color, justify="left"); timer_label.place(relx=0.01, rely=1.0, anchor="sw", y=-15)
    external_clock_label = tk.Label(message_window, text="", font=("Helvetica", 48, "bold"), fg=selected_fg_color, bg=selected_bg_color); external_clock_label.place(relx=0.99, rely=1.0, anchor="se", y=-15)
    update_external_clock(external_clock_label)
    secili_hastane_obj = next((h for h in load_data(HASTANELER_DOSYASI, []) if h.get('ad') == selected_hospital.get()), None)
    if secili_hastane_obj:
        hospital_name_label_on_screen.config(text=turkce_upper(secili_hastane_obj.get('ad', '')))
        try:
            img = Image.open(secili_hastane_obj['logo']); original_width, original_height = img.size
            if original_height > 0: aspect_ratio = original_width / original_height; new_width = int(200 * aspect_ratio); resized_img = img.resize((new_width, 200), Image.Resampling.LANCZOS); active_logo_ref = ImageTk.PhotoImage(resized_img); tk.Label(logo_frame, image=active_logo_ref, bg=selected_bg_color).pack(pady=0)
        except Exception as e: print(f"Logo yükleme hatası: {e}")
    main_message = text_area.get("1.0", tk.END).strip()
    if main_message and not main_message.startswith("Notunuzu buraya"):
        g_main_message_label = tk.Label(main_content_frame, text=main_message, font=("Helvetica", 63, "bold"), fg=selected_fg_color, bg=selected_bg_color, wraplength=chosen_monitor.width - 150, justify="center"); g_main_message_label.pack(pady=(20, 10))
    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    doctor_name = turkce_upper(secilen_profil_obj.get('icerik', '')) if secilen_profil_obj else ""
    if doctor_name: tk.Label(main_content_frame, text=doctor_name, font=("Helvetica", 48, "normal"), fg=selected_fg_color, bg=selected_bg_color, wraplength=chosen_monitor.width - 100).pack(pady=10)
    message_window.bind("<Escape>", lambda e: close_secondary_screen()); 
    update_all_button_states()
def update_preview():
    global active_preview_logo_ref;
    if not preview_square_frame: return
    for widget in preview_square_frame.winfo_children(): widget.destroy()
    preview_inner_frame = tk.Frame(preview_square_frame, bg=selected_bg_color); preview_inner_frame.pack(expand=True, fill="both")
    secili_hastane_obj = next((h for h in load_data(HASTANELER_DOSYASI, []) if h.get('ad') == selected_hospital.get()), None)
    if secili_hastane_obj:
        tk.Label(preview_inner_frame, text=turkce_upper(secili_hastane_obj.get('ad', '')), font=("Helvetica", 10, "bold"), fg=selected_fg_color, bg=selected_bg_color, wraplength=PREVIEW_SIZE-20).pack(pady=2)
        try:
            img = Image.open(secili_hastane_obj['logo']); img.thumbnail((80, 40), Image.Resampling.LANCZOS); active_preview_logo_ref = ImageTk.PhotoImage(img); tk.Label(preview_inner_frame, image=active_preview_logo_ref, bg=selected_bg_color).pack(pady=2)
        except Exception: pass
    main_message = text_area.get("1.0", tk.END).strip()
    if main_message and not main_message.startswith("Notunuzu buraya"): tk.Label(preview_inner_frame, text=main_message, font=("Helvetica", 12, "bold"), fg=selected_fg_color, bg=selected_bg_color, wraplength=PREVIEW_SIZE-20, justify="center").pack(pady=5, expand=True)
    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    doctor_name = turkce_upper(secilen_profil_obj.get('icerik', '')) if secilen_profil_obj else ""
    if doctor_name: tk.Label(preview_inner_frame, text=doctor_name, font=("Helvetica", 10, "normal"), fg=selected_fg_color, bg=selected_bg_color).pack(pady=2)
def start_stop_countdown():
    global countdown_seconds_remaining, is_countdown_paused, countdown_job_id, start_stop_button, is_warning_sent, is_on_break, is_on_stopwatch_break, stopwatch_seconds
    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    doktor_adi = secilen_profil_obj.get('ad', 'Doktor') if secilen_profil_obj else 'Doktor'
    if is_countdown_paused:
        is_countdown_paused = False
        if countdown_seconds_remaining is not None:
            mins, secs = divmod(countdown_seconds_remaining, 60)
            send_telegram_message(f"{doktor_adi} mola sayacını devam ettirdi. Kalan süre: {mins:02d}:{secs:02d}", notification_type='mola')
        run_countdown(); update_all_button_states(); return
    if is_on_break and not is_countdown_paused:
        is_countdown_paused = True
        if countdown_job_id: root.after_cancel(countdown_job_id); countdown_job_id = None
        try:
            total_mins, total_secs = map(int, return_time_entry.get().strip().split(':')); toplam_saniye = (total_mins * 60) + total_secs
            kullanilan_saniye = toplam_saniye - countdown_seconds_remaining; kull_mins, kull_secs = divmod(kullanilan_saniye, 60)
            kal_mins, kal_secs = divmod(countdown_seconds_remaining, 60)
            send_telegram_message(f"{doktor_adi} mola sayacını durdurdu.\nKullanılan Süre: {kull_mins:02d}:{kull_secs:02d}\nKalan Süre: {kal_mins:02d}:{kal_secs:02d}", notification_type='mola')
        except Exception as e: print(f"Durdurma mesajı için süre hesaplama hatası: {e}")
        update_all_button_states(); return
    mola_suresi_str = return_time_entry.get().strip()
    if secilen_profil_obj:
        mola_mesaji_adi = secilen_profil_obj.get("mola_mesaji_adi")
        if mola_mesaji_adi:
            mola_mesaji_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == mola_mesaji_adi), None)
            if mola_mesaji_obj: set_message_text(mola_mesaji_obj.get('icerik', '')); show_message()
    if mola_suresi_str:
        is_on_break = True
        try:
            minutes, seconds = map(int, mola_suresi_str.split(':')); countdown_seconds_remaining = (minutes * 60) + seconds
            if countdown_seconds_remaining <= 0: messagebox.showwarning("Geçersiz Süre", "Mola süresi sıfırdan büyük olmalıdır."); is_on_break = False; update_all_button_states(); return
            is_countdown_paused, is_warning_sent = False, False
            send_telegram_message(f"{doktor_adi} molaya çıktı. Mola süresi: {mola_suresi_str}", notification_type='mola')
            log_mola_activity(doktor_adi, "Mola Başladı", "Geri Sayım", planlanan_sure=mola_suresi_str); run_countdown()
        except (ValueError, AttributeError): messagebox.showwarning("Geçersiz Süre", "Lütfen mola süresini 'dd:ss' formatında girin."); is_on_break = False
    else:
        is_on_stopwatch_break = True; stopwatch_seconds = 0
        send_telegram_message(f"{doktor_adi} süresiz bir molaya çıktı.", notification_type='mola')
        log_mola_activity(doktor_adi, "Mola Başladı", "Kronometre"); run_stopwatch_break()
    update_all_button_states()
def reset_countdown():
    global countdown_seconds_remaining, is_countdown_paused, countdown_job_id, is_warning_sent, is_on_break, is_on_stopwatch_break, stopwatch_job_id, stopwatch_seconds, is_on_overtime_break, overtime_job_id, overtime_seconds
    
    if is_on_break or is_on_stopwatch_break or is_on_overtime_break:
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
        doktor_adi = secilen_profil_obj.get('ad', 'Doktor') if secilen_profil_obj else 'Doktor'
        mola_suresi_str = return_time_entry.get().strip() or pip_mola_entry.get().strip()

        if is_on_overtime_break:
            try:
                total_mins, total_secs = map(int, mola_suresi_str.split(':')); toplam_saniye = (total_mins * 60) + total_secs
                gerceklesen_toplam_saniye = toplam_saniye + overtime_seconds
                g_mins, g_secs = divmod(int(gerceklesen_toplam_saniye), 60); gerceklesen_sure_str = f"{g_mins:02d}:{g_secs:02d}"
                send_telegram_message(f"{doktor_adi}, molasını bitirdi. Toplam süre: {gerceklesen_sure_str}.", notification_type='mola')
                log_mola_activity(doktor_adi, "Mola Manuel Bitirildi", "Geri Sayım (Uzatmalı)", planlanan_sure=mola_suresi_str, gerceklesen_sure=gerceklesen_sure_str)
            except Exception as e:
                 print(f"Uzatmalı mola loglama hatası: {e}")
                 send_telegram_message(f"{doktor_adi} molasını bitirdi.", notification_type='mola')
                 log_mola_activity(doktor_adi, "Mola Manuel Bitirildi", "Geri Sayım (Uzatmalı)")
        elif is_on_break and countdown_seconds_remaining is not None:
            try:
                total_mins, total_secs = map(int, mola_suresi_str.split(':')); toplam_saniye = (total_mins * 60) + total_secs
                kullanilan_saniye = toplam_saniye - countdown_seconds_remaining
                if not (0 <= kullanilan_saniye <= toplam_saniye): kullanilan_saniye = toplam_saniye
                kull_mins, kull_secs = divmod(int(kullanilan_saniye), 60); kullanilan_sure_str = f"{kull_mins:02d}:{kull_secs:02d}"
                send_telegram_message(f"{doktor_adi}, molasını bitirdi/sıfırladı. Geçen süre: {kullanilan_sure_str}.", notification_type='mola')
                log_mola_activity(doktor_adi, "Mola Manuel Bitirildi", "Geri Sayım", planlanan_sure=mola_suresi_str, gerceklesen_sure=kullanilan_sure_str)
            except Exception:
                send_telegram_message(f"{doktor_adi} molasını bitirdi.", notification_type='mola')
                log_mola_activity(doktor_adi, "Mola Manuel Bitirildi", "Geri Sayım")
        elif is_on_stopwatch_break:
             gecen_sure_str = time.strftime('%M:%S', time.gmtime(stopwatch_seconds))
             send_telegram_message(f"{doktor_adi} süresiz molasını bitirdi. Geçen süre: {gecen_sure_str}.", notification_type='mola')
             log_mola_activity(doktor_adi, "Mola Manuel Bitirildi", "Kronometre", gerceklesen_sure=gecen_sure_str)

    if countdown_job_id: root.after_cancel(countdown_job_id); countdown_job_id = None
    if stopwatch_job_id: root.after_cancel(stopwatch_job_id); stopwatch_job_id = None
    if overtime_job_id: root.after_cancel(overtime_job_id); overtime_job_id = None
    
    is_on_break, is_on_stopwatch_break, stopwatch_seconds, is_countdown_paused, is_warning_sent = False, False, 0, False, False
    is_on_overtime_break, overtime_seconds = False, 0
    countdown_seconds_remaining = None

    if timer_label and timer_label.winfo_exists(): timer_label.config(text="")
    if local_timer_label and local_timer_label.winfo_exists(): local_timer_label.config(text="")
    if pip_mola_button and pip_mola_button.winfo_exists(): pip_mola_button.config(text="Mola")
    if pip_mola_entry and pip_mola_entry.winfo_exists(): pip_mola_entry.delete(0, tk.END)

    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
    if secilen_profil_obj:
        mola_sonrasi_mesaji_adi = secilen_profil_obj.get("mola_sonrasi_mesaji_adi")
        if mola_sonrasi_mesaji_adi:
            mola_sonrasi_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == mola_sonrasi_mesaji_adi), None)
            if mola_sonrasi_obj:
                new_content = mola_sonrasi_obj.get('icerik', '')
                set_message_text(new_content)
                if message_window and message_window.winfo_exists(): show_message()
                else: update_preview()
    update_all_button_states(); update_pip_button_states()

def run_countdown():
    global countdown_seconds_remaining, is_countdown_paused, countdown_job_id, is_warning_sent, g_main_message_label, is_on_overtime_break, overtime_seconds, overtime_job_id
    if is_countdown_paused or countdown_seconds_remaining is None: return
    if countdown_seconds_remaining >= 0:
        try:
            uyari_suresi_str = warning_time_entry.get().strip()
            if not is_warning_sent and uyari_suresi_str:
                warn_mins, warn_secs = map(int, uyari_suresi_str.split(':')); uyari_saniyesi = (warn_mins * 60) + warn_secs
                if countdown_seconds_remaining == uyari_saniyesi:
                    secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
                    doktor_adi = secilen_profil_obj.get('ad', 'Doktor') if secilen_profil_obj else 'Doktor'
                    send_telegram_message(f"{doktor_adi}'un molasının bitmesine kalan süre: {uyari_suresi_str}", notification_type='mola'); is_warning_sent = True
        except (ValueError, AttributeError): pass
        mins, secs = divmod(countdown_seconds_remaining, 60); time_str = '{:02d}:{:02d}'.format(mins, secs)
        if timer_label and timer_label.winfo_exists(): timer_label.config(text=f"Mola\n{time_str}")
        if local_timer_label and local_timer_label.winfo_exists(): local_timer_label.config(text=time_str)
        if pip_mola_button and pip_mola_button.winfo_exists(): pip_mola_button.config(text=time_str)
        countdown_seconds_remaining -= 1; countdown_job_id = root.after(1000, run_countdown)
    else:
        is_on_overtime_break = True; overtime_seconds = 0
        
        secilen_profil_obj = next((p for p in load_data(PROFILER_DOSYASI, []) if p.get('ad') == selected_profile.get()), None)
        doktor_adi = secilen_profil_obj.get('ad', 'Doktor') if secilen_profil_obj else 'Doktor'
        send_telegram_message(f"{doktor_adi}'un planlanan molası bitti, sayaç uzatmaları sayıyor.", notification_type='mola')
        
        planlanan_sure = return_time_entry.get().strip()
        log_mola_activity(doktor_adi, "Mola Tamamlandı", "Geri Sayım", planlanan_sure=planlanan_sure, gerceklesen_sure=planlanan_sure)
        
        if secilen_profil_obj:
            mola_sonrasi_mesaji_adi = secilen_profil_obj.get("mola_sonrasi_mesaji_adi")
            if mola_sonrasi_mesaji_adi:
                mola_sonrasi_obj = next((m for m in load_data(MESAJLAR_DOSYASI, []) if m.get('ad') == mola_sonrasi_mesaji_adi), None)
                if mola_sonrasi_obj:
                    new_content = mola_sonrasi_obj.get('icerik', '')
                    set_message_text(new_content)
                    if g_main_message_label and g_main_message_label.winfo_exists(): g_main_message_label.config(text=new_content)
                    if timer_label and timer_label.winfo_exists(): timer_label.config(text="")
        
        update_all_button_states()
        run_overtime_break_timer()

def run_overtime_break_timer():
    global overtime_seconds, overtime_job_id, is_on_overtime_break
    if not is_on_overtime_break: return
    
    mins, secs = divmod(overtime_seconds, 60)
    time_str = f"+{mins:02d}:{secs:02d}"
    
    if local_timer_label and local_timer_label.winfo_exists(): local_timer_label.config(text=time_str)
    if pip_mola_button and pip_mola_button.winfo_exists(): pip_mola_button.config(text=time_str)
    
    overtime_seconds += 1
    overtime_job_id = root.after(1000, run_overtime_break_timer)

def update_external_clock(clock_widget):
    global external_clock_job_id
    if clock_widget and clock_widget.winfo_exists(): clock_widget.config(text=time.strftime('%H:%M:%S')); external_clock_job_id = root.after(1000, update_external_clock, clock_widget)
    elif 'external_clock_job_id' in globals() and external_clock_job_id: root.after_cancel(external_clock_job_id); external_clock_job_id = None

def update_all_button_states(*args):
    is_screen_open = message_window and message_window.winfo_exists()
    any_break_active = is_on_break or is_on_stopwatch_break or is_on_overtime_break
    is_action_active = any_break_active or is_patient_inside
    
    start_stop_state = 'disabled'
    if is_screen_open and not is_action_active:
        start_stop_state = 'normal'
    if is_on_break:
        start_stop_state = 'normal'

    end_reset_state = 'normal' if any_break_active else 'disabled'
    patient_call_state = 'normal' if is_screen_open and not is_action_active else 'disabled'
    
    if start_stop_button:
        button_text = "Başlat"
        if is_on_break:
            button_text = "Durdur" if not is_countdown_paused else "Devam Et"
        start_stop_button.config(state=start_stop_state, text=button_text)
    
    if end_break_button: end_break_button.config(state=end_reset_state)
    if reset_button: reset_button.config(state=end_reset_state)
    if return_time_entry: return_time_entry.config(state= 'disabled' if is_action_active else 'normal')
    if call_patient_btn: call_patient_btn.config(state=patient_call_state)
    if reset_patient_btn: reset_patient_btn.config(state=patient_call_state)
    if patient_entry: patient_entry.config(state=patient_call_state)
    if patient_combobox: patient_combobox.config(state=patient_call_state)
    if refresh_patient_btn: refresh_patient_btn.config(state=patient_call_state)
    if patient_inside_button: patient_inside_button.config(state='normal' if is_screen_open and not any_break_active else 'disabled')
    if close_screen_button: close_screen_button.config(state='normal')
    update_pip_button_states()

def update_pip_button_states():
    if not (pip_window and pip_window.winfo_exists()): return
    any_break_active = is_on_break or is_on_stopwatch_break or is_on_overtime_break
    mola_state = 'disabled' if is_patient_inside else 'normal'
    hasta_state = 'disabled' if any_break_active else 'normal'
    
    if pip_mola_button: pip_mola_button.config(state=mola_state)
    if pip_mola_entry: pip_mola_entry.config(state=mola_state)
    if pip_patient_button: pip_patient_button.config(state=hasta_state)
    
    if 'pip_patient_combobox' in globals() and pip_patient_combobox and pip_patient_combobox.winfo_exists():
        pip_patient_combobox.config(state=hasta_state)
        pip_refresh_patient_btn.config(state=hasta_state)
        pip_call_patient_btn.config(state=hasta_state)
        pip_reset_patient_btn.config(state=hasta_state)
        
    if 'pip_screen_button' in globals() and pip_screen_button and pip_screen_button.winfo_exists():
        if message_window and message_window.winfo_exists():
            pip_screen_button.config(text="Ekranı Kapat", style="Red.PiP.TButton")
        else:
            pip_screen_button.config(text="Ekranı Aç", style="Green.PiP.TButton")
            
    if 'pip_main_window_button' in globals() and pip_main_window_button and pip_main_window_button.winfo_exists():
        if root and root.winfo_viewable():
            pip_main_window_button.config(text="Ana Paneli Gizle", style="Yellow.PiP.TButton")
        else:
            pip_main_window_button.config(text="Ana Paneli Aç", style="Blue.TButton")

def hide_window():
    if root: root.withdraw()
    update_pip_button_states()
    
def show_window():
    global root
    if not root or not root.winfo_exists(): root = tk.Tk(); setup_gui(root)
    else: root.deiconify(); root.lift()
    update_pip_button_states()
    
def exit_action(tray_icon):
    if tray_icon: tray_icon.stop()
    if root: root.destroy()
    
def close_secondary_screen(keep_timer=False):
    global countdown_job_id, external_clock_job_id, overtime_job_id
    if not keep_timer:
        if countdown_job_id: root.after_cancel(countdown_job_id); countdown_job_id = None
        if overtime_job_id: root.after_cancel(overtime_job_id); overtime_job_id = None

    if 'external_clock_job_id' in globals() and external_clock_job_id:
        try: root.after_cancel(external_clock_job_id)
        except: pass
        external_clock_job_id = None
    if message_window and message_window.winfo_exists(): message_window.destroy()
    if screen_menu: screen_menu.config(style="TMenubutton")
    if close_screen_button: close_screen_button.config(text="Dış Ekran Kapalı", style="Green.TButton")
    update_all_button_states()

def setup_gui(master_window):
    global root, text_area, return_time_entry, bg_color_swatch, fg_color_swatch, selected_screen, preview_square_frame, patient_entry, patient_combobox, refresh_patient_btn, selected_profile, profile_menu, selected_message, message_menu, selected_hospital, hospital_menu, local_timer_label, start_stop_button, reset_button, mola_telegram_var, hasta_telegram_var, mola_telegram_check, hasta_telegram_check, screen_menu, warning_time_entry, close_screen_button, patient_inside_button, patient_inside_timer_entry, call_patient_btn, reset_patient_btn, end_break_button, selected_patient_var, log_mola_var, log_cagirma_var, log_muayene_var, log_notion_mola_var, log_notion_cagirma_var, log_notion_muayene_var, log_notion_mola_check, log_notion_cagirma_check, log_notion_muayene_check, mask_patient_name_var
    root = master_window; root.title("Poliklinik Ekran Yöneticisi v6.3"); root.geometry("1500x740"); root.minsize(1500, 740)
    style = ttk.Style(); style.theme_use('clam')
    style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 9, 'bold'))
    style.configure("Blue.TButton", background="#007bff", foreground="white"); style.map("Blue.TButton", background=[('active', '#0056b3')])
    style.configure("Green.TButton", background="#28a745", foreground="white"); style.map("Green.TButton", background=[('active', '#218838')])
    style.configure("Red.TButton", background="#dc3545", foreground="white"); style.map("Red.TButton", background=[('active', '#c82333')])
    style.configure("Yellow.TButton", background="#ffc107", foreground="black"); style.map("Yellow.TButton", background=[('active', '#e0a800')])
    style.configure("Switch.TCheckbutton", font=('Helvetica', 9, 'bold')); style.configure("White.TMenubutton", background="white"); style.configure("TLabelFrame.Label", font=("Helvetica", 11, "bold"))
    
    main_pane = ttk.PanedWindow(root, orient='horizontal'); main_pane.pack(fill='both', expand=True, padx=5, pady=5)
    controls_container = ttk.Frame(main_pane, padding=5); main_pane.add(controls_container, weight=1)
    canvas = tk.Canvas(controls_container, highlightthickness=0); scrollbar = ttk.Scrollbar(controls_container, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas); scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw"); canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y"); scrollable_frame.columnconfigure(1, weight=1)
    
    def create_profile_row(parent, row, label_text, var, manager_command):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky='w', padx=5, pady=2)
        menu = ttk.OptionMenu(parent, var, ""); menu.grid(row=row, column=1, sticky='ew', padx=5)
        ttk.Button(parent, text="...", width=3, command=manager_command).grid(row=row, column=2, sticky='e', padx=5)
        return menu
    
    selected_hospital = tk.StringVar(root); selected_profile = tk.StringVar(root); selected_message = tk.StringVar(root); selected_patient_var = tk.StringVar(root)
    mesaj_listesi = load_data(MESAJLAR_DOSYASI, []); mesaj_adlari = [m.get('ad') for m in mesaj_listesi]
    profile_cols = {'ad': 'Görünen Ad', 'icerik': 'İçerik (Tam Ad)', 'telegram_token': 'Telegram Bot Token', 'telegram_chat_id': 'Telegram Chat ID', 'api_url': 'API URL', 'api_bearer': 'API Bearer (Poliklinik)', 'notion_token': 'Notion API Anahtarı', 'notion_mola_db_id': 'Notion Mola DB ID', 'notion_cagirma_db_id': 'Notion Çağırma DB ID', 'notion_muayene_db_id': 'Notion Muayene DB ID', 'mola_mesaji_adi': 'Varsayılan Mola Mesajı', 'mola_sonrasi_mesaji_adi': 'Mola Sonrası Mesajı', 'hasta_var_mesaji_adi': 'Varsayılan Hasta Var Mesajı'}
    profile_dropdowns = {'mola_mesaji_adi': mesaj_adlari, 'mola_sonrasi_mesaji_adi': mesaj_adlari, 'hasta_var_mesaji_adi': mesaj_adlari}
    hospital_menu = create_profile_row(scrollable_frame, 0, "Hastane Profili:", selected_hospital, lambda: open_data_manager("Hastane Yöneticisi", HASTANELER_DOSYASI, {'ad': 'Hastane Adı', 'logo': 'Logo Dosyası'}, refresh_hospital_dropdown, VARSAYILAN_HASTANELER, True))
    profile_menu = create_profile_row(scrollable_frame, 1, "Doktor Profili:", selected_profile, lambda: open_data_manager("Profil Yöneticisi", PROFILER_DOSYASI, profile_cols, refresh_profile_dropdown, VARSAYILAN_PROFILER, dropdown_fields=profile_dropdowns))
    message_menu = create_profile_row(scrollable_frame, 2, "Hazır Mesaj:", selected_message, lambda: open_data_manager("Mesaj Yöneticisi", MESAJLAR_DOSYASI, {'ad': 'Görünen Ad', 'icerik': 'Mesaj İçeriği'}, refresh_message_dropdown, VARSAYILAN_MESAJLAR))
    selected_hospital.trace_add("write", lambda *args: update_preview()); selected_profile.trace_add("write", profile_changed); selected_message.trace_add("write", message_changed)
    text_area = scrolledtext.ScrolledText(scrollable_frame, wrap=tk.WORD, font=("Helvetica", 11), height=8); text_area.grid(row=3, column=0, columnspan=3, sticky='ew', pady=5, padx=5); text_area.bind("<KeyRelease>", lambda e: update_preview())
    settings_frame = ttk.Frame(scrollable_frame); settings_frame.grid(row=4, column=0, columnspan=3, sticky='ew', padx=5, pady=5); settings_frame.columnconfigure(1, weight=1)
    ttk.Label(settings_frame, text="Hedef Ekran:").grid(row=0, column=0, sticky='w')
    screen_options = [f"Ekran {i+1} ({m.width}x{m.height})" for i, m in enumerate(get_monitors())]; selected_screen = tk.StringVar(root)
    if screen_options: selected_screen.set(screen_options[0] if len(screen_options) == 1 else screen_options[1] if len(screen_options) > 1 else "")
    screen_menu = ttk.OptionMenu(settings_frame, selected_screen, *([selected_screen.get()] + screen_options)); screen_menu.grid(row=0, column=1, sticky='ew', padx=5)
    color_frame = ttk.Frame(settings_frame); color_frame.grid(row=0, column=2, sticky='e')
    bg_color_swatch = tk.Label(color_frame, text="Fon", relief='raised'); bg_color_swatch.pack(side='left', padx=2, ipady=2, ipadx=5); bg_color_swatch.bind("<Button-1>", lambda e: choose_color(bg_color_swatch, 'bg'))
    fg_color_swatch = tk.Label(color_frame, text="Yazı", relief='raised'); fg_color_swatch.pack(side='left', padx=2, ipady=2, ipadx=5); fg_color_swatch.bind("<Button-1>", lambda e: choose_color(fg_color_swatch, 'fg'))
    patient_inside_frame = ttk.LabelFrame(scrollable_frame, text="Muayene Süresi (Plan)", padding=5); patient_inside_frame.grid(row=5, column=0, columnspan=3, sticky='ew', pady=(10, 2), padx=5); patient_inside_frame.columnconfigure(1, weight=1)
    patient_inside_timer_entry = ttk.Entry(patient_inside_frame, width=10, font=("Helvetica", 10)); patient_inside_timer_entry.pack(side='left', padx=(0, 5))
    patient_inside_button = ttk.Button(patient_inside_frame, text="Hasta Yok", style="Green.TButton", command=toggle_patient_inside); patient_inside_button.pack(side='left', fill='x', expand=True, padx=(0, 5))
    pip_toggle_button = ttk.Button(patient_inside_frame, text="PiP", style="Blue.TButton", command=toggle_pip_window, width=5); pip_toggle_button.pack(side='right')

    log_settings_frame = ttk.LabelFrame(scrollable_frame, text="Loglama Ayarları", padding=10)
    log_settings_frame.grid(row=6, column=0, columnspan=3, sticky='ew', pady=5, padx=5)
    log_mola_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(log_settings_frame, text="Mola Sürelerini Excel'e Logla", variable=log_mola_var, style="Switch.TCheckbutton").pack(anchor='w', fill='x')
    log_cagirma_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(log_settings_frame, text="Hasta Çağırmayı Excel'e Logla", variable=log_cagirma_var, style="Switch.TCheckbutton").pack(anchor='w', fill='x')
    log_muayene_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(log_settings_frame, text="Muayene Sürelerini Excel'e Logla", variable=log_muayene_var, style="Switch.TCheckbutton").pack(anchor='w', fill='x')
    ttk.Separator(log_settings_frame, orient='horizontal').pack(fill='x', pady=5)
    log_notion_mola_var = tk.BooleanVar(value=True)
    log_notion_mola_check = ttk.Checkbutton(log_settings_frame, text="Mola Sürelerini Notion'a Logla", variable=log_notion_mola_var, style="Switch.TCheckbutton", state='disabled')
    log_notion_mola_check.pack(anchor='w', fill='x')
    log_notion_cagirma_var = tk.BooleanVar(value=True)
    log_notion_cagirma_check = ttk.Checkbutton(log_settings_frame, text="Hasta Çağırmayı Notion'a Logla", variable=log_notion_cagirma_var, style="Switch.TCheckbutton", state='disabled')
    log_notion_cagirma_check.pack(anchor='w', fill='x')
    log_notion_muayene_var = tk.BooleanVar(value=True)
    log_notion_muayene_check = ttk.Checkbutton(log_settings_frame, text="Muayene Sürelerini Notion'a Logla", variable=log_notion_muayene_var, style="Switch.TCheckbutton", state='disabled')
    log_notion_muayene_check.pack(anchor='w', fill='x')

    show_message_btn_frame = ttk.Frame(scrollable_frame); show_message_btn_frame.grid(row=7, column=0, columnspan=3, sticky='ew', pady=5, padx=5)
    ttk.Button(show_message_btn_frame, text="MESAJI GÖSTER", style="Blue.TButton", command=show_message).pack(fill='x', expand=True, ipady=8)
    action_buttons_frame = ttk.Frame(scrollable_frame); action_buttons_frame.grid(row=8, column=0, columnspan=3, sticky='ew', pady=5, padx=5)
    ttk.Button(action_buttons_frame, text="Temizle", style="Green.TButton", command=clear_form).pack(side='left', expand=True, fill='x', padx=2)
    close_screen_button = ttk.Button(action_buttons_frame, text="Dış Ekran Kapalı", style="Green.TButton", command=lambda: close_secondary_screen(keep_timer=False)); close_screen_button.pack(side='left', expand=True, fill='x', padx=2)
    ttk.Button(action_buttons_frame, text="Kapat", style="Red.TButton", command=lambda: exit_action(tray_icon)).pack(side='left', expand=True, fill='x', padx=2)
    
    right_pane = ttk.PanedWindow(main_pane, orient='vertical')
    main_pane.add(right_pane, weight=1)
    main_pane.sashpos(0, 684)

    preview_container = ttk.LabelFrame(right_pane, text="Canlı Görüntü", padding=10)
    right_pane.add(preview_container, weight=1)
    
    preview_square_frame = tk.Frame(preview_container, width=PREVIEW_SIZE, height=PREVIEW_SIZE, relief="sunken", borderwidth=1)
    preview_square_frame.pack(expand=True, fill='both')
    preview_square_frame.pack_propagate(False)

    actions_container = ttk.Frame(right_pane)
    right_pane.add(actions_container, weight=1)
    
    timer_controls_frame = ttk.LabelFrame(actions_container, text="Mola & Sayaç", padding=10); timer_controls_frame.pack(fill='x', expand=False, pady=(0, 5)); timer_controls_frame.columnconfigure(1, weight=1)
    ttk.Label(timer_controls_frame, text="Mola (dd:ss):").grid(row=0, column=0, sticky="w"); return_time_entry = ttk.Entry(timer_controls_frame, font=("Helvetica", 10), width=10); return_time_entry.grid(row=0, column=1, sticky="ew", padx=5);
    ttk.Label(timer_controls_frame, text="Uyarı (dd:ss):").grid(row=1, column=0, sticky="w", pady=(5,0)); warning_time_entry = ttk.Entry(timer_controls_frame, font=("Helvetica", 10), width=10); warning_time_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=(5,0))
    timer_btn_frame = ttk.Frame(timer_controls_frame); timer_btn_frame.grid(row=2, column=0, columnspan=2, pady=(10,0))
    start_stop_button = ttk.Button(timer_btn_frame, text="Başlat", style="Green.TButton", command=start_stop_countdown); start_stop_button.pack(side="left", expand=True, fill="x", padx=(0,5))
    end_break_button = ttk.Button(timer_btn_frame, text="Molayı Bitir", style="Yellow.TButton", command=reset_countdown); end_break_button.pack(side="left", expand=True, fill="x", padx=(0,5))
    reset_button = ttk.Button(timer_btn_frame, text="Sıfırla", style="Red.TButton", command=reset_countdown); reset_button.pack(side="left", expand=True, fill="x")
    local_timer_label = ttk.Label(timer_controls_frame, text="", font=("Helvetica", 28, "bold"), anchor="center"); local_timer_label.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
    mola_telegram_var = tk.BooleanVar(value=False); mola_telegram_check = ttk.Checkbutton(timer_controls_frame, text="Mola Telegram Bildirimleri", variable=mola_telegram_var, style="Switch.TCheckbutton", state="disabled"); mola_telegram_check.grid(row=4, column=0, columnspan=2, pady=(5,0), sticky='w')
    
    patient_call_frame_preview = ttk.LabelFrame(actions_container, text="Hasta Çağırma", padding=10); patient_call_frame_preview.pack(fill='x', expand=False); patient_call_frame_preview.columnconfigure(0, weight=1)
    patient_entry = ttk.Entry(patient_call_frame_preview, font=("Helvetica", 11))
    patient_combobox = ttk.Combobox(patient_call_frame_preview, font=("Helvetica", 11), textvariable=selected_patient_var)
    refresh_patient_btn = ttk.Button(patient_call_frame_preview, text="🔄", width=3, command=fetch_and_populate_patients, style="Blue.TButton"); patient_entry.grid(row=0, column=0, sticky='ew')
    button_frame = ttk.Frame(patient_call_frame_preview); button_frame.grid(row=0, column=2, sticky='e')
    call_patient_btn = ttk.Button(button_frame, text="Çağır", style="Red.TButton", command=call_patient); call_patient_btn.pack(side='left', padx=(5,2))
    reset_patient_btn = ttk.Button(button_frame, text="Sıfırla", style="Red.TButton", command=reset_patient_call); reset_patient_btn.pack(side='left')
    hasta_telegram_var = tk.BooleanVar(value=False); hasta_telegram_check = ttk.Checkbutton(patient_call_frame_preview, text="Telegram Bildirimi Gönder", variable=hasta_telegram_var, style="Switch.TCheckbutton", state="disabled"); hasta_telegram_check.grid(row=1, column=0, columnspan=3, pady=(5,0), sticky='w')
    
    mask_patient_name_var = tk.BooleanVar(value=True) # Varsayılan olarak isimler maskeli gelsin
    mask_patient_name_check = ttk.Checkbutton(patient_call_frame_preview, text="Hasta İsimlerini Maskele (Örn: AH*** YIL***)", variable=mask_patient_name_var, style="Switch.TCheckbutton")
    mask_patient_name_check.grid(row=2, column=0, columnspan=3, pady=(5,0), sticky='w')

    refresh_hospital_dropdown(); refresh_profile_dropdown(); refresh_message_dropdown()
    if load_data(PROFILER_DOSYASI, []): profile_changed()
    if load_data(MESAJLAR_DOSYASI, []): message_changed()
    else: on_text_area_focus_out(None)
    update_preview(); 
    update_all_button_states()

def main():
    global tray_icon, root
    if not os.path.exists("hastaneler"): os.makedirs("hastaneler")
    load_data(HASTANELER_DOSYASI, VARSAYILAN_HASTANELER); load_data(PROFILER_DOSYASI, VARSAYILAN_PROFILER); load_data(MESAJLAR_DOSYASI, VARSAYILAN_MESAJLAR)
    try: image = Image.open("ikon.ico")
    except FileNotFoundError:
        img_data=b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAjpJREFUOE+dk1tIk1E1xz/zzGN33M7tTtt0Sy62dENL4aSFLKyUhUIUHxZ5UESfQS8i6EHQHZKgiHoRBV1EXyQkSGuGMu2a5GtaupLz0Wbd3T2dzYx7fvA4gocWfMLf3fvw/J7fL38C8J/v0c2nS+ST42CTeACuB5PAprYtRPzKSWF+BbwG7oCTsLwF8LguH0zgLxiGs2AXeAf83kL8BfA8uB7Mhkl8C+A1sB2cBGphF4+eBTyH/wMXgZfAQqAe+AxcB6bDPDDHB74bAHYfusAs8P12WzYg2AVugG2wFvAnQJOgCvweuAXsA/vAOYHw82GAL2A7uA/sghvB6yLkM2ADuAQ2wR8U+DswCuwE6sC7wTvgCVgG5oAd4CfwPLgGfA0mQI+AHeB94BhwFdgBvoGfL3wHvoN3wTfgHfgOfA4eAM/gM3gKvAfvgVfgB/gYfAHeAz+BP+B18A+cAj4Bf4GXwFvgN/gAfAE+g4/AV+At8B+cAy4CD4DfA9eAD8B54CfYhGcl4JvgObAFnAE/At+AW8CmwAdgDdgCvgVPgR/A7eA1cBW4DjrgETgJvAs8B74Gj4C54DPwO/ANeAj8Ab4DnoB3wPfgYfAqeA48B74Cn4BvwJvgZfAC+Ap8D94CHwHvgC/Ab+B58AR4Dnxi08A18B54FjgD/gPfgV/gGfAJ+Ap8AF4EvgNvgUfAWeD3eBp8BV4CHwOvgE/gYfAWeAz8Bp4CXwIvgZfA78Df4CHwDfgAfAE+ge8hC8d3YBe4Ywz4A0wCKzHMA/gC3hMmw+s3WJ+G/wEgm45/iL83kL8AAAAASUVORK5CYII='
        image = Image.open(io.BytesIO(base64.b64decode(img_data)))
    menu = (item('Kontrol Panelini Aç', show_window, default=True), item('Çıkış', lambda: exit_action(tray_icon)))
    tray_icon = icon("Poliklinik Ekran Yöneticisi", image, "Poliklinik Ekran Yöneticisi", menu)
    threading.Thread(target=tray_icon.run, daemon=True).start()
    show_window(); tk.mainloop()

if __name__ == "__main__":
    try: main()
    except Exception as e:
        error_message = f"Beklenmedik bir hata oluştu:\n\n{e}"; traceback_details = traceback.format_exc()
        try: hata_penceresi = tk.Tk(); hata_penceresi.withdraw(); messagebox.showerror("Kritik Hata", error_message); hata_penceresi.destroy()
        except: pass
        with open("hata_gunlugu.txt", "w", encoding="utf-8") as f:
            f.write(f"Hata Tarihi: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n"); f.write(error_message); f.write("\n\nTeknik Detaylar:\n"); f.write(traceback_details)