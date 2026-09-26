"""
Plansculpt Auto Mailer Pro - Enterprise Bulk Outreach Suite
Author: Plansculpt & S. S. M Jahir Jahan Khan Miru
"""

import os
import sys
import json
import time
import re
import datetime
import threading
import smtplib
import calendar
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formatdate
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image

import pandas as pd
import customtkinter as ctk

# HTML Rendering engine import with graceful fallback
try:
    from tkinterweb import HtmlFrame
    HAS_TKINTERWEB = True
except ImportError:
    HAS_TKINTERWEB = False

try:
    import tkhtmlview
    HAS_TKHTMLVIEW = True
except ImportError:
    HAS_TKHTMLVIEW = False

# --- Set Appearance & Theme ---
ctk.set_appearance_mode("Dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

CONFIG_FILE = "config.json"
QUEUE_FILE = "scheduled_queue.json"
DEFAULT_TEMPLATE_FILE = "email_template.html"
LOGO_PATH = os.path.join("Assets", "logo.jpg")

SMTP_PRESETS = {
    "Gmail": {"host": "smtp.gmail.com", "port": 587, "tls": True},
    "Outlook / Office 365": {"host": "smtp.office365.com", "port": 587, "tls": True},
    "Yahoo": {"host": "smtp.mail.yahoo.com", "port": 587, "tls": True},
    "Custom SMTP": {"host": "", "port": 587, "tls": True}
}

import zoneinfo

BD_TIMEZONE = zoneinfo.ZoneInfo("Asia/Dhaka")

TIMEZONE_MAP = {
    # --- USA TIMEZONES (Prioritized at Top) ---
    "[USA] US Eastern Time (ET - New York, Boston, DC, Atlanta)": "America/New_York",
    "[USA] US Central Time (CT - Chicago, Dallas, Houston)": "America/Chicago",
    "[USA] US Mountain Time (MT - Denver, Phoenix, Salt Lake City)": "America/Denver",
    "[USA] US Pacific Time (PT - Los Angeles, San Francisco, Seattle)": "America/Los_Angeles",
    "[USA] US Alaska Time (AKT - Anchorage)": "America/Anchorage",
    "[USA] US Hawaii Time (HST - Honolulu)": "Pacific/Honolulu",
    
    # --- ALL OTHER TIMEZONES (Sorted chronologically by UTC offset: UTC-12 to UTC+14) ---
    "UTC-12:00 (Baker Island, Howland Island)": "Etc/GMT+12",
    "UTC-11:00 (American Samoa, Niue)": "Pacific/Pago_Pago",
    "UTC-10:00 (Hawaii-Aleutian / Tahiti)": "Pacific/Honolulu",
    "UTC-09:00 (Alaska / Gambier)": "America/Anchorage",
    "UTC-08:00 (Pacific Time - Los Angeles, Vancouver)": "America/Los_Angeles",
    "UTC-07:00 (Mountain Time - Denver, Phoenix, Calgary)": "America/Denver",
    "UTC-06:00 (Central Time - Chicago, Mexico City)": "America/Chicago",
    "UTC-05:00 (Eastern Time - New York, Toronto, Miami, Lima)": "America/New_York",
    "UTC-04:00 (Atlantic Time - Halifax, Santiago, Caracas)": "America/Halifax",
    "UTC-03:30 (Newfoundland - St. John's)": "America/St_Johns",
    "UTC-03:00 (Buenos Aires, Sao Paulo, Montevideo)": "America/Argentina/Buenos_Aires",
    "UTC-02:00 (Mid-Atlantic / Fernando de Noronha)": "America/Noronha",
    "UTC-01:00 (Azores, Cape Verde)": "Atlantic/Azores",
    "UTC+00:00 (GMT / UTC - London, Dublin, Lisbon, Accra, Reykjavik)": "Europe/London",
    "UTC+01:00 (CET - Berlin, Paris, Rome, Madrid, Amsterdam)": "Europe/Paris",
    "UTC+02:00 (EET - Athens, Cairo, Helsinki, Kyiv, Jerusalem, Beirut)": "Europe/Athens",
    "UTC+03:00 (MSK / AST - Moscow, Riyadh, Istanbul, Nairobi, Doha)": "Europe/Moscow",
    "UTC+03:30 (IRST - Iran / Tehran)": "Asia/Tehran",
    "UTC+04:00 (GST - Dubai, Abu Dhabi, Baku, Tbilisi, Yerevan)": "Asia/Dubai",
    "UTC+04:30 (AFT - Afghanistan / Kabul)": "Asia/Kabul",
    "UTC+05:00 (PKT / UZT - Karachi, Islamabad, Tashkent, Yekaterinburg)": "Asia/Karachi",
    "UTC+05:30 (IST - India / New Delhi, Mumbai, Kolkata, Colombo)": "Asia/Kolkata",
    "UTC+05:45 (NPT - Nepal / Kathmandu)": "Asia/Kathmandu",
    "UTC+06:00 (BST - Bangladesh Standard Time / Dhaka, Almaty)": "Asia/Dhaka",
    "UTC+06:30 (MMT - Myanmar / Yangon)": "Asia/Yangon",
    "UTC+07:00 (ICT - Bangkok, Hanoi, Jakarta, Novosibirsk)": "Asia/Bangkok",
    "UTC+08:00 (CST / SGT - Singapore, Beijing, Hong Kong, Taipei, Perth)": "Asia/Singapore",
    "UTC+08:45 (ACWST - Eucla)": "Australia/Eucla",
    "UTC+09:00 (JST / KST - Tokyo, Seoul, Osaka)": "Asia/Tokyo",
    "UTC+09:30 (ACST - Adelaide, Darwin)": "Australia/Adelaide",
    "UTC+10:00 (AEST - Sydney, Melbourne, Brisbane, Vladivostok)": "Australia/Sydney",
    "UTC+10:30 (Lord Howe Island)": "Australia/Lord_Howe",
    "UTC+11:00 (SBT / AEDT - Solomon Islands, New Caledonia)": "Pacific/Guadalcanal",
    "UTC+12:00 (NZST / FJT - Auckland, Wellington, Fiji)": "Pacific/Auckland",
    "UTC+12:45 (CHAST - Chatham Islands)": "Pacific/Chatham",
    "UTC+13:00 (TOT - Tonga, Samoa, Tokelau)": "Pacific/Tongatapu",
    "UTC+14:00 (LINT - Line Islands, Kiritimati)": "Pacific/Kiritimati"
}

TIMEZONE_PRESETS = list(TIMEZONE_MAP.keys())

def get_timezone_obj(tz_label):
    iana_name = TIMEZONE_MAP.get(tz_label, "America/New_York")
    try:
        return zoneinfo.ZoneInfo(iana_name)
    except Exception:
        return zoneinfo.ZoneInfo("UTC")

def format_date_display(date_input, ref_today=None):
    if ref_today is None:
        ref_today = datetime.datetime.now().date()
    try:
        if isinstance(date_input, datetime.datetime):
            d = date_input.date()
        elif isinstance(date_input, datetime.date):
            d = date_input
        else:
            m = re.search(r'(\d{4}-\d{2}-\d{2})', str(date_input))
            if m:
                d = datetime.datetime.strptime(m.group(1), "%Y-%m-%d").date()
            else:
                return str(date_input)
        
        diff = (d - ref_today).days
        if diff == 0:
            tag = "Today"
        elif diff == 1:
            tag = "Tomorrow"
        elif diff == -1:
            tag = "Yesterday"
        else:
            tag = d.strftime("%a")
        return f"{d.strftime('%Y-%m-%d')} ({tag})"
    except Exception:
        return str(date_input)

class CalendarPickerModal(ctk.CTkToplevel):
    """Modern Dark-Themed Interactive Calendar Modal with Past-Date Disabling"""
    def __init__(self, parent, initial_date=None, min_date=None, title="Select Date", callback=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x420")
        self.resizable(False, False)
        self.callback = callback
        
        self.today = datetime.datetime.now().date()
        self.min_date = min_date if min_date else self.today

        if initial_date:
            if isinstance(initial_date, str):
                m = re.search(r'(\d{4}-\d{2}-\d{2})', initial_date)
                if m:
                    self.selected_date = datetime.datetime.strptime(m.group(1), "%Y-%m-%d").date()
                else:
                    self.selected_date = self.today
            elif isinstance(initial_date, (datetime.date, datetime.datetime)):
                self.selected_date = initial_date.date() if isinstance(initial_date, datetime.datetime) else initial_date
            else:
                self.selected_date = self.today
        else:
            self.selected_date = self.today

        self.view_year = self.selected_date.year
        self.view_month = self.selected_date.month

        self.protocol("WM_DELETE_WINDOW", self.close_modal)

        self.create_widgets()
        self.render_calendar()

        # Center on parent & bring to focus
        try:
            p_x = parent.winfo_rootx()
            p_y = parent.winfo_rooty()
            p_w = parent.winfo_width()
            p_h = parent.winfo_height()
            x = p_x + (p_w - 350) // 2
            y = p_y + (p_h - 420) // 2
            self.geometry(f"350x420+{max(50, x)}+{max(50, y)}")
            self.transient(parent)
            self.lift()
            self.focus_force()
            self.after(60, lambda: self.grab_set() if self.winfo_exists() else None)
        except Exception:
            pass

    def create_widgets(self):
        self.main_box = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=10)
        self.main_box.pack(fill="both", expand=True, padx=8, pady=8)

        # Nav Header
        nav_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        nav_frame.pack(fill="x", padx=10, pady=(10, 6))

        self.btn_prev = ctk.CTkButton(nav_frame, text="◀", width=36, height=28, fg_color="#334155", hover_color="#475569", command=self.prev_month)
        self.btn_prev.pack(side="left")

        self.lbl_month_year = ctk.CTkLabel(nav_frame, text="", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        self.lbl_month_year.pack(side="left", expand=True)

        self.btn_next = ctk.CTkButton(nav_frame, text="▶", width=36, height=28, fg_color="#334155", hover_color="#475569", command=self.next_month)
        self.btn_next.pack(side="right")

        # Weekday headers
        days_header_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        days_header_frame.pack(fill="x", padx=10, pady=2)
        for col, day_name in enumerate(["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]):
            days_header_frame.grid_columnconfigure(col, weight=1)
            lbl = ctk.CTkLabel(
                days_header_frame, 
                text=day_name, 
                font=ctk.CTkFont(size=11, weight="bold"), 
                text_color="#38bdf8" if col in (0, 6) else "#94a3b8"
            )
            lbl.grid(row=0, column=col, padx=1, pady=2, sticky="ew")

        # Grid container
        self.days_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        self.days_frame.pack(fill="both", expand=True, padx=10, pady=4)
        for c in range(7):
            self.days_frame.grid_columnconfigure(c, weight=1)

        # Action row
        action_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        action_frame.pack(fill="x", padx=10, pady=(6, 10))

        btn_today = ctk.CTkButton(
            action_frame, 
            text="Today", 
            width=70, 
            height=28, 
            fg_color="#334155", 
            hover_color="#475569", 
            font=ctk.CTkFont(size=11, weight="bold"), 
            command=lambda: self.select_day(self.today)
        )
        btn_today.pack(side="left", padx=2)

        btn_tomorrow = ctk.CTkButton(
            action_frame, 
            text="Tomorrow", 
            width=80, 
            height=28, 
            fg_color="#334155", 
            hover_color="#475569", 
            font=ctk.CTkFont(size=11, weight="bold"), 
            command=lambda: self.select_day(self.today + datetime.timedelta(days=1))
        )
        btn_tomorrow.pack(side="left", padx=2)

        btn_cancel = ctk.CTkButton(
            action_frame, 
            text="Cancel", 
            width=65, 
            height=28, 
            fg_color="#dc2626", 
            hover_color="#b91c1c", 
            font=ctk.CTkFont(size=11, weight="bold"), 
            command=self.close_modal
        )
        btn_cancel.pack(side="right", padx=2)

    def prev_month(self):
        if self.view_month == 1:
            self.view_month = 12
            self.view_year -= 1
        else:
            self.view_month -= 1
        self.render_calendar()

    def next_month(self):
        if self.view_month == 12:
            self.view_month = 1
            self.view_year += 1
        else:
            self.view_month += 1
        self.render_calendar()

    def render_calendar(self):
        for widget in self.days_frame.winfo_children():
            widget.destroy()

        month_name = calendar.month_name[self.view_month]
        self.lbl_month_year.configure(text=f"{month_name} {self.view_year}")

        if self.min_date and (self.view_year < self.min_date.year or (self.view_year == self.min_date.year and self.view_month <= self.min_date.month)):
            self.btn_prev.configure(state="disabled")
        else:
            self.btn_prev.configure(state="normal")

        cal = calendar.Calendar(firstweekday=6)
        month_dates = list(cal.itermonthdates(self.view_year, self.view_month))

        for idx, d in enumerate(month_dates):
            r = idx // 7
            c = idx % 7
            self.days_frame.grid_rowconfigure(r, weight=1)

            is_curr_month = (d.month == self.view_month)
            is_past = bool(self.min_date and d < self.min_date)
            is_selected = (d == self.selected_date)
            is_today = (d == self.today)

            if is_past or not is_curr_month:
                btn_state = "disabled"
                fg = "transparent"
                txt_color = "#475569" if is_past else "#64748b"
                do_hover = False
                hover_col = None
            else:
                btn_state = "normal"
                do_hover = True
                if is_selected:
                    fg = "#0284c7"
                    txt_color = "white"
                    hover_col = "#0369a1"
                else:
                    fg = "#1e293b"
                    txt_color = "white"
                    hover_col = "#334155"

            btn_kwargs = {
                "text": str(d.day),
                "width": 34,
                "height": 30,
                "corner_radius": 6,
                "fg_color": fg,
                "text_color": txt_color,
                "font": ctk.CTkFont(size=11, weight="bold" if (is_selected or is_today) else "normal"),
                "border_width": 1 if is_today and not is_selected else 0,
                "state": btn_state,
                "hover": do_hover,
                "command": lambda day_val=d: self.select_day(day_val)
            }
            if is_today and not is_selected:
                btn_kwargs["border_color"] = "#38bdf8"
            if hover_col:
                btn_kwargs["hover_color"] = hover_col

            btn = ctk.CTkButton(self.days_frame, **btn_kwargs)
            btn.grid(row=r, column=c, padx=2, pady=2, sticky="nsew")

    def select_day(self, date_obj):
        if self.min_date and date_obj < self.min_date:
            return
        self.selected_date = date_obj
        if self.callback:
            self.callback(date_obj.strftime("%Y-%m-%d"))
        self.close_modal()

    def close_modal(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

class TimezoneSearchModal(ctk.CTkToplevel):
    """Modern Searchable & Scrollable Timezone Modal with Fast Filtering and Region Chips"""
    def __init__(self, parent, timezone_list, selected_tz=None, callback=None):
        super().__init__(parent)
        self.title("Select Target Region / Timezone")
        self.geometry("540x580")
        self.resizable(False, False)
        self.callback = callback
        self.timezone_list = timezone_list
        self.selected_tz = selected_tz or (timezone_list[0] if timezone_list else "")
        self.current_category = "All"

        self.protocol("WM_DELETE_WINDOW", self.close_modal)

        self.create_widgets()
        self.filter_timezones()

        # Center on parent & bring to focus
        try:
            p_x = parent.winfo_rootx()
            p_y = parent.winfo_rooty()
            p_w = parent.winfo_width()
            p_h = parent.winfo_height()
            x = p_x + (p_w - 540) // 2
            y = p_y + (p_h - 580) // 2
            self.geometry(f"540x580+{max(50, x)}+{max(50, y)}")
            self.transient(parent)
            self.lift()
            self.focus_force()
            self.after(60, lambda: self.grab_set() if self.winfo_exists() else None)
        except Exception:
            pass

    def create_widgets(self):
        self.main_box = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=10)
        self.main_box.pack(fill="both", expand=True, padx=8, pady=8)

        # Header Title
        lbl_title = ctk.CTkLabel(
            self.main_box, 
            text="🌍 Select Target Timezone & Region", 
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#38bdf8"
        )
        lbl_title.pack(anchor="w", padx=14, pady=(12, 4))

        # Search Bar
        search_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        search_frame.pack(fill="x", padx=12, pady=(4, 6))

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.filter_timezones())
        
        self.search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="🔍 Search country, city, code (e.g. USA, New York, London, Dhaka, UTC+6)...",
            height=34,
            font=ctk.CTkFont(size=12)
        )
        self.search_entry.pack(fill="x")

        # Filter Chips
        chips_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        chips_frame.pack(fill="x", padx=12, pady=(0, 6))

        self.chip_buttons = {}
        categories = [("All", "All"), ("USA", "🇺🇸 USA"), ("Europe", "🇪🇺 Europe"), ("Asia", "🌏 Asia"), ("Americas", "🌎 Americas"), ("Pacific", "🌊 Pacific")]
        for key, label in categories:
            btn = ctk.CTkButton(
                chips_frame,
                text=label,
                height=26,
                width=65,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#0284c7" if key == "All" else "#1e293b",
                hover_color="#0369a1",
                command=lambda k=key: self.set_category(k)
            )
            btn.pack(side="left", padx=2)
            self.chip_buttons[key] = btn

        # Scrollable list
        self.scroll_list = ctk.CTkScrollableFrame(self.main_box, fg_color="#090d16", corner_radius=8)
        self.scroll_list.pack(fill="both", expand=True, padx=12, pady=(2, 8))

        # Bottom Action Bar
        bottom_frame = ctk.CTkFrame(self.main_box, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=12, pady=(4, 8))

        self.lbl_count = ctk.CTkLabel(bottom_frame, text="", font=ctk.CTkFont(size=11), text_color="#94a3b8")
        self.lbl_count.pack(side="left")

        btn_close = ctk.CTkButton(
            bottom_frame,
            text="Close",
            width=75,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            command=self.close_modal
        )
        btn_close.pack(side="right")

    def set_category(self, cat_key):
        self.current_category = cat_key
        for k, btn in self.chip_buttons.items():
            btn.configure(fg_color="#0284c7" if k == cat_key else "#1e293b")
        self.filter_timezones()

    def filter_timezones(self):
        for w in self.scroll_list.winfo_children():
            w.destroy()

        query = self.search_var.get().strip().lower()
        matched = []

        for tz in self.timezone_list:
            tz_lower = tz.lower()
            # Category match
            if self.current_category == "USA" and not ("[usa]" in tz_lower or "united states" in tz_lower):
                continue
            elif self.current_category == "Europe" and not ("london" in tz_lower or "paris" in tz_lower or "berlin" in tz_lower or "rome" in tz_lower or "athens" in tz_lower or "cet" in tz_lower or "eet" in tz_lower or "gmt" in tz_lower or "moscow" in tz_lower):
                continue
            elif self.current_category == "Asia" and not ("dhaka" in tz_lower or "tokyo" in tz_lower or "singapore" in tz_lower or "beijing" in tz_lower or "india" in tz_lower or "karachi" in tz_lower or "dubai" in tz_lower or "bangkok" in tz_lower or "seoul" in tz_lower or "bst" in tz_lower or "ist" in tz_lower or "cst" in tz_lower or "jst" in tz_lower):
                continue
            elif self.current_category == "Americas" and not ("[usa]" in tz_lower or "toronto" in tz_lower or "buenos aires" in tz_lower or "sao paulo" in tz_lower or "halifax" in tz_lower or "mexico" in tz_lower or "chile" in tz_lower or "atlantic" in tz_lower):
                continue
            elif self.current_category == "Pacific" and not ("sydney" in tz_lower or "auckland" in tz_lower or "hawaii" in tz_lower or "fiji" in tz_lower or "honolulu" in tz_lower or "samoa" in tz_lower or "nzst" in tz_lower or "aest" in tz_lower):
                continue

            # Query match
            if query and query not in tz_lower:
                continue

            matched.append(tz)

        self.lbl_count.configure(text=f"Showing {len(matched)} of {len(self.timezone_list)} timezones")

        if not matched:
            lbl_empty = ctk.CTkLabel(
                self.scroll_list,
                text="No matching timezones found.\nTry a different city or offset keyword.",
                text_color="#94a3b8",
                font=ctk.CTkFont(size=12)
            )
            lbl_empty.pack(pady=30)
            return

        for tz in matched:
            is_selected = (tz == self.selected_tz)
            is_usa = "[USA]" in tz

            card = ctk.CTkFrame(
                self.scroll_list,
                fg_color="#0284c7" if is_selected else ("#1e293b" if is_usa else "#131d2e"),
                corner_radius=6,
                border_width=1 if is_selected or is_usa else 0
            )
            if is_selected:
                card.configure(border_color="#38bdf8")
            elif is_usa:
                card.configure(border_color="#3b82f6")
            card.pack(fill="x", pady=2, padx=2)

            btn = ctk.CTkButton(
                card,
                text=tz,
                anchor="w",
                height=32,
                fg_color="transparent",
                hover_color="#334155" if not is_selected else "#0369a1",
                text_color="white" if is_selected else ("#38bdf8" if is_usa else "#e2e8f0"),
                font=ctk.CTkFont(size=11, weight="bold" if (is_selected or is_usa) else "normal"),
                command=lambda val=tz: self.select_timezone(val)
            )
            btn.pack(fill="x", padx=4, pady=2)

    def select_timezone(self, tz_val):
        self.selected_tz = tz_val
        if self.callback:
            self.callback(tz_val)
        self.close_modal()

    def close_modal(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

class AutoMailerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Window Setup ---
        self.title("Plansculpt Auto Mailer Pro")
        self.geometry("1320x850")
        self.minsize(1120, 700)

        # State Variables
        self.recipients_df = None
        self.data_filepath = ""
        self.attachment_paths = []
        self.is_sending = False
        self.is_paused = False
        self.should_stop = False
        self.campaign_thread = None
        self.scheduler_worker_thread = None
        self.scheduler_active = True
        self.sent_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        self.total_targets = 0
        self.campaign_log_records = []
        self._live_preview_after_id = None
        self._syntax_highlight_after_id = None
        self.editor_font_size = 12
        self._syncing_timezones = False

        # Scheduled Batches Queue (Persistent)
        self.scheduled_batches = self.load_scheduled_queue()

        # Load configuration
        self.config_data = self.load_config()

        # Build UI
        self.setup_grid_layout()
        self.create_sidebar()
        self.create_main_content()

        # Load initial defaults
        self.load_initial_defaults()

        # Start Background Scheduler Engine
        self.start_background_scheduler_loop()

    # ==========================================
    #             CONFIG & PERSISTENCE
    # ==========================================
    def load_config(self):
        default_config = {
            "email": "",
            "password": "",
            "smtp_provider": "Gmail",
            "smtp_host": "smtp.gmail.com",
            "smtp_port": 587,
            "smtp_tls": True,
            "sender_name": "",
            "default_subject": "",
            "delay_seconds": 12,
            "random_jitter": 4,
            "batch_pause_count": 20,
            "batch_pause_minutes": 5,
            "daily_limit": 150,
            "last_attachment": "Jahir_Jahan_Miru_CV.pdf" if os.path.exists("Jahir_Jahan_Miru_CV.pdf") else "",
            "editor_font_size": 12,
            "appearance_mode": "Dark"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    default_config.update(loaded)
            except Exception as e:
                print(f"Error loading config: {e}")
        self.editor_font_size = default_config.get("editor_font_size", 12)
        return default_config

    def save_config(self):
        try:
            self.config_data["email"] = self.smtp_email_entry.get().strip()
            self.config_data["password"] = self.smtp_pass_entry.get().strip()
            self.config_data["sender_name"] = self.sender_name_entry.get().strip()
            self.config_data["smtp_provider"] = self.smtp_provider_var.get()
            self.config_data["smtp_host"] = self.smtp_host_entry.get().strip()
            self.config_data["smtp_port"] = int(self.smtp_port_entry.get().strip() or 587)
            self.config_data["smtp_tls"] = self.smtp_tls_var.get()
            self.config_data["default_subject"] = self.subject_entry.get().strip()
            self.config_data["delay_seconds"] = int(self.delay_spinbox.get() or 12)
            self.config_data["random_jitter"] = int(self.jitter_spinbox.get() or 4)
            self.config_data["batch_pause_count"] = int(self.batch_count_spinbox.get() or 20)
            self.config_data["batch_pause_minutes"] = int(self.batch_pause_spinbox.get() or 5)
            self.config_data["daily_limit"] = int(self.daily_limit_spinbox.get() or 150)
            self.config_data["editor_font_size"] = self.editor_font_size
            self.config_data["appearance_mode"] = ctk.get_appearance_mode()
            
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config_data, f, indent=4)
        except Exception as e:
            self.log_message(f"[-] Error saving settings: {e}", "error")

    def load_scheduled_queue(self):
        if os.path.exists(QUEUE_FILE):
            try:
                with open(QUEUE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return []
        return []

    def save_scheduled_queue(self):
        try:
            with open(QUEUE_FILE, "w", encoding="utf-8") as f:
                json.dump(self.scheduled_batches, f, indent=4)
        except Exception as e:
            print(f"Error saving queue: {e}")

    # ==========================================
    #                LAYOUT SETUP
    # ==========================================
    def setup_grid_layout(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def create_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=245, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(9, weight=1)

        # Brand / Logo Header
        brand_container = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        brand_container.grid(row=0, column=0, padx=16, pady=(18, 12), sticky="ew")

        # Load Logo Image if available
        if os.path.exists(LOGO_PATH):
            try:
                pil_logo = Image.open(LOGO_PATH)
                self.logo_ctk = ctk.CTkImage(light_image=pil_logo, dark_image=pil_logo, size=(42, 42))
                logo_lbl = ctk.CTkLabel(brand_container, image=self.logo_ctk, text="")
                logo_lbl.pack(side="left", padx=(0, 10))
            except Exception as e:
                print(f"Error loading logo: {e}")

        text_brand_frame = ctk.CTkFrame(brand_container, fg_color="transparent")
        text_brand_frame.pack(side="left", fill="both", expand=True)

        self.logo_label = ctk.CTkLabel(
            text_brand_frame, 
            text="Plansculpt", 
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#38bdf8",
            anchor="w"
        )
        self.logo_label.pack(fill="x")

        self.sub_logo_label = ctk.CTkLabel(
            text_brand_frame, 
            text="Auto Mailer Pro", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray80",
            anchor="w"
        )
        self.sub_logo_label.pack(fill="x")

        # Subtle separator
        sep = ctk.CTkFrame(self.sidebar_frame, height=1, fg_color=("gray75", "gray30"))
        sep.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")

        # Navigation Buttons (Clean English, Professional Alignment)
        self.nav_buttons = {}
        pages = [
            ("Campaign Monitor", "campaign"),
            ("Recipients & Data", "data"),
            ("Template Editor", "template"),
            ("Campaign Scheduler", "schedule"),
            ("Anti-Spam & Delivery", "antispam"),
            ("SMTP Settings", "settings"),
            ("Activity Logs", "logs")
        ]

        for idx, (title, page_id) in enumerate(pages, start=2):
            btn = ctk.CTkButton(
                self.sidebar_frame, 
                text=title,
                anchor="w",
                height=36,
                font=ctk.CTkFont(size=13, weight="normal"),
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray75", "gray25"),
                command=lambda p=page_id: self.show_page(p)
            )
            btn.grid(row=idx, column=0, padx=12, pady=3, sticky="ew")
            self.nav_buttons[page_id] = btn

        # Theme Selector at bottom
        self.theme_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Theme:", font=ctk.CTkFont(size=11), text_color="gray")
        self.theme_label.grid(row=10, column=0, padx=16, pady=(10, 2), sticky="w")

        self.theme_mode_menu = ctk.CTkOptionMenu(
            self.sidebar_frame, 
            values=["Dark", "Light", "System"],
            command=self.change_appearance_mode
        )
        self.theme_mode_menu.grid(row=11, column=0, padx=16, pady=(0, 16), sticky="ew")
        self.theme_mode_menu.set(self.config_data.get("appearance_mode", "Dark"))

    def create_main_content(self):
        self.pages_container = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.pages_container.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)
        self.pages_container.grid_rowconfigure(0, weight=1)
        self.pages_container.grid_columnconfigure(0, weight=1)

        self.pages = {}
        self.create_campaign_page()
        self.create_data_page()
        self.create_template_page()
        self.create_schedule_page()
        self.create_antispam_page()
        self.create_settings_page()
        self.create_logs_page()

        # Show initial page
        self.show_page("campaign")

    def show_page(self, page_name):
        for name, frame in self.pages.items():
            frame.grid_forget()
            if name in self.nav_buttons:
                self.nav_buttons[name].configure(fg_color="transparent", font=ctk.CTkFont(size=13, weight="normal"))

        if page_name in self.pages:
            self.pages[page_name].grid(row=0, column=0, sticky="nsew")
            if page_name in self.nav_buttons:
                self.nav_buttons[page_name].configure(
                    fg_color=("gray70", "gray28"), 
                    font=ctk.CTkFont(size=13, weight="bold")
                )
            if page_name == "template":
                self.trigger_live_preview_update()
                self.apply_syntax_highlighting()
            if page_name == "schedule":
                self.refresh_scheduled_batches_table()

    def change_appearance_mode(self, new_mode):
        ctk.set_appearance_mode(new_mode)
        self.save_config()
        self.trigger_live_preview_update()

    # ==========================================
    #           PAGE 1: CAMPAIGN & SENDER
    # ==========================================
    def create_campaign_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["campaign"] = page
        page.grid_rowconfigure(3, weight=1)
        page.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Header Cards
        self.card_total = self.create_stat_card(page, 0, 0, "Total Targets", "0", "#3b82f6")
        self.card_sent = self.create_stat_card(page, 0, 1, "Sent Successfully", "0", "#10b981")
        self.card_failed = self.create_stat_card(page, 0, 2, "Failed", "0", "#ef4444")
        self.card_status = self.create_stat_card(page, 0, 3, "Campaign Status", "IDLE", "#6366f1")

        # Controls Box
        controls_box = ctk.CTkFrame(page)
        controls_box.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(15, 10), padx=5)
        controls_box.grid_columnconfigure(1, weight=1)

        subj_lbl = ctk.CTkLabel(controls_box, text="Email Subject:", font=ctk.CTkFont(weight="bold"))
        subj_lbl.grid(row=0, column=0, padx=15, pady=(12, 6), sticky="w")

        self.subject_entry = ctk.CTkEntry(
            controls_box, 
            placeholder_text="Write your subject here (supports placeholders like {Professor_Name}, {Institute_Name})..."
        )
        self.subject_entry.grid(row=0, column=1, columnspan=2, padx=15, pady=(12, 6), sticky="ew")

        attach_lbl = ctk.CTkLabel(controls_box, text="Attachments:", font=ctk.CTkFont(weight="bold"))
        attach_lbl.grid(row=1, column=0, padx=15, pady=6, sticky="w")

        attach_frame = ctk.CTkFrame(controls_box, fg_color="transparent")
        attach_frame.grid(row=1, column=1, columnspan=2, padx=15, pady=6, sticky="ew")
        attach_frame.grid_columnconfigure(0, weight=1)

        self.attachments_label = ctk.CTkLabel(
            attach_frame, 
            text="No files attached (Optional: PDF, DOCX, CV, Portfolio)", 
            anchor="w",
            text_color="gray"
        )
        self.attachments_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        btn_attach = ctk.CTkButton(
            attach_frame, 
            text="Add File(s)", 
            width=100, 
            command=self.browse_attachments
        )
        btn_attach.grid(row=0, column=1, padx=2)

        btn_clear_attach = ctk.CTkButton(
            attach_frame, 
            text="Clear", 
            width=65, 
            fg_color="#dc2626", 
            hover_color="#b91c1c",
            command=self.clear_attachments
        )
        btn_clear_attach.grid(row=0, column=2, padx=2)

        # Action Buttons Row
        actions_frame = ctk.CTkFrame(page, fg_color="transparent")
        actions_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=10, padx=5)
        actions_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self.btn_start = ctk.CTkButton(
            actions_frame, 
            text="Start Immediate Send", 
            height=40, 
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#10b981", 
            hover_color="#059669",
            command=self.start_campaign
        )
        self.btn_start.grid(row=0, column=0, padx=4, sticky="ew")

        self.btn_pause = ctk.CTkButton(
            actions_frame, 
            text="Pause", 
            height=40, 
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#f59e0b", 
            hover_color="#d97706",
            state="disabled",
            command=self.toggle_pause
        )
        self.btn_pause.grid(row=0, column=1, padx=4, sticky="ew")

        self.btn_stop = ctk.CTkButton(
            actions_frame, 
            text="Stop", 
            height=40, 
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#ef4444", 
            hover_color="#dc2626",
            state="disabled",
            command=self.stop_campaign
        )
        self.btn_stop.grid(row=0, column=2, padx=4, sticky="ew")

        self.btn_test = ctk.CTkButton(
            actions_frame, 
            text="Send Test Email", 
            height=40, 
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#6366f1", 
            hover_color="#4f46e5",
            command=self.show_test_email_dialog
        )
        self.btn_test.grid(row=0, column=3, padx=4, sticky="ew")

        self.btn_preview = ctk.CTkButton(
            actions_frame, 
            text="Preview Sample", 
            height=40, 
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#0284c7", 
            hover_color="#0369a1",
            command=self.preview_sample_email
        )
        self.btn_preview.grid(row=0, column=4, padx=4, sticky="ew")

        # Progress Box
        progress_box = ctk.CTkFrame(page)
        progress_box.grid(row=3, column=0, columnspan=4, sticky="nsew", pady=(5, 5), padx=5)
        progress_box.grid_rowconfigure(2, weight=1)
        progress_box.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(progress_box, height=14)
        self.progress_bar.grid(row=0, column=0, padx=15, pady=(15, 6), sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_box, 
            text="Ready. Select data file and start campaign or schedule for automated delivery.", 
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        self.progress_label.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")

        self.quick_log_text = ctk.CTkTextbox(
            progress_box, 
            font=ctk.CTkFont(family="Consolas", size=11), 
            wrap="word"
        )
        self.quick_log_text.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="nsew")

    def create_stat_card(self, parent, row, col, title, value, color_hex):
        card = ctk.CTkFrame(parent, corner_radius=8, fg_color=("gray85", "gray17"))
        card.grid(row=row, column=col, padx=4, pady=4, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)

        t_lbl = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12, weight="bold"), text_color="gray")
        t_lbl.grid(row=0, column=0, padx=12, pady=(10, 2), sticky="w")

        v_lbl = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=22, weight="bold"), text_color=color_hex)
        v_lbl.grid(row=1, column=0, padx=12, pady=(0, 10), sticky="w")
        return v_lbl

    # ==========================================
    #           PAGE 2: DATA & RECIPIENTS
    # ==========================================
    def create_data_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["data"] = page
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)

        top_bar = ctk.CTkFrame(page)
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_bar.grid_columnconfigure(1, weight=1)

        file_lbl = ctk.CTkLabel(top_bar, text="Recipients File:", font=ctk.CTkFont(weight="bold"))
        file_lbl.grid(row=0, column=0, padx=12, pady=10, sticky="w")

        self.data_path_entry = ctk.CTkEntry(top_bar, placeholder_text="Select a CSV or Excel (.xlsx, .xls) file...")
        self.data_path_entry.grid(row=0, column=1, padx=8, pady=10, sticky="ew")

        btn_browse = ctk.CTkButton(top_bar, text="Browse File", width=110, command=self.browse_recipients_file)
        btn_browse.grid(row=0, column=2, padx=(4, 8), pady=10)

        btn_reload = ctk.CTkButton(
            top_bar, 
            text="🔄 Reload Data", 
            width=120, 
            fg_color="#0284c7", 
            hover_color="#0369a1",
            font=ctk.CTkFont(weight="bold"),
            command=self.reload_current_recipients_file
        )
        btn_reload.grid(row=0, column=3, padx=(0, 12), pady=10)

        info_bar = ctk.CTkFrame(page)
        info_bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        info_bar.grid_columnconfigure(3, weight=1)

        self.data_summary_label = ctk.CTkLabel(
            info_bar, 
            text="No file loaded. Supported columns: 'Professor Name', 'Institute Name', 'Email Address', 'Recent Paper Topics'",
            font=ctk.CTkFont(size=12)
        )
        self.data_summary_label.grid(row=0, column=0, padx=12, pady=8, sticky="w")

        table_container = ctk.CTkFrame(page)
        table_container.grid(row=2, column=0, sticky="nsew")
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview", 
            background="#1e293b", 
            foreground="white", 
            fieldbackground="#1e293b", 
            rowheight=26,
            font=("Segoe UI", 10)
        )
        style.configure("Custom.Treeview.Heading", background="#334155", foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Custom.Treeview", background=[('selected', '#2563eb')])

        self.data_tree = ttk.Treeview(table_container, style="Custom.Treeview", show="headings", selectmode="extended")
        self.data_tree.bind("<Double-1>", lambda event: self.edit_selected_recipient())
        
        tree_scroll_y = ttk.Scrollbar(table_container, orient="vertical", command=self.data_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_container, orient="horizontal", command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        self.data_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")

        bottom_tools = ctk.CTkFrame(page, fg_color="transparent")
        bottom_tools.grid(row=3, column=0, sticky="ew", pady=(10, 0))

        btn_add = ctk.CTkButton(
            bottom_tools, 
            text="+ Add Contact", 
            width=115,
            font=ctk.CTkFont(weight="bold"),
            fg_color="#0f766e",
            hover_color="#0d9488",
            command=self.add_new_recipient
        )
        btn_add.grid(row=0, column=0, padx=4, sticky="w")

        btn_edit = ctk.CTkButton(
            bottom_tools, 
            text="Edit Row", 
            width=90,
            font=ctk.CTkFont(weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self.edit_selected_recipient
        )
        btn_edit.grid(row=0, column=1, padx=4, sticky="w")

        btn_del = ctk.CTkButton(
            bottom_tools, 
            text="Delete Row", 
            width=95,
            font=ctk.CTkFont(weight="bold"),
            fg_color="#b91c1c",
            hover_color="#991b1b",
            command=self.delete_selected_recipient
        )
        btn_del.grid(row=0, column=2, padx=4, sticky="w")

        btn_dedup = ctk.CTkButton(
            bottom_tools, 
            text="Remove Duplicates", 
            fg_color="gray30",
            hover_color="gray40",
            command=self.deduplicate_recipients
        )
        btn_dedup.grid(row=0, column=3, padx=4, sticky="w")

        btn_clean_invalid = ctk.CTkButton(
            bottom_tools, 
            text="Clean Invalid / Empty", 
            fg_color="gray30",
            hover_color="gray40",
            command=self.remove_invalid_recipients
        )
        btn_clean_invalid.grid(row=0, column=4, padx=4, sticky="w")

        btn_save_csv = ctk.CTkButton(
            bottom_tools,
            text="Save / Export CSV",
            font=ctk.CTkFont(weight="bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            command=self.save_current_recipients_file
        )
        btn_save_csv.grid(row=0, column=5, padx=(10, 4), sticky="w")

    # ==========================================
    #   PAGE 3: TEMPLATE EDITOR (VSCODE STYLED)
    # ==========================================
    def create_template_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["template"] = page
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)

        # Top Bar
        top_ctrl_bar = ctk.CTkFrame(page)
        top_ctrl_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top_ctrl_bar.grid_columnconfigure(3, weight=1)

        mode_lbl = ctk.CTkLabel(top_ctrl_bar, text="View Mode:", font=ctk.CTkFont(weight="bold"))
        mode_lbl.grid(row=0, column=0, padx=(10, 4), pady=8, sticky="w")

        self.view_mode_seg = ctk.CTkSegmentedButton(
            top_ctrl_bar,
            values=["Split (Both)", "Editor Only", "Preview Only"],
            command=self.on_view_mode_changed
        )
        self.view_mode_seg.grid(row=0, column=1, padx=4, pady=8, sticky="w")
        self.view_mode_seg.set("Split (Both)")

        # Zoom Controls
        zoom_frame = ctk.CTkFrame(top_ctrl_bar, fg_color="transparent")
        zoom_frame.grid(row=0, column=2, padx=8, pady=8, sticky="w")

        btn_zoom_out = ctk.CTkButton(zoom_frame, text="A-", width=34, height=28, font=ctk.CTkFont(weight="bold"), fg_color="gray30", hover_color="gray40", command=self.zoom_out)
        btn_zoom_out.pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(zoom_frame, text=f"{int((self.editor_font_size/12)*100)}%", width=46, font=ctk.CTkFont(size=11, weight="bold"))
        self.lbl_zoom.pack(side="left", padx=2)

        btn_zoom_in = ctk.CTkButton(zoom_frame, text="A+", width=34, height=28, font=ctk.CTkFont(weight="bold"), fg_color="gray30", hover_color="gray40", command=self.zoom_in)
        btn_zoom_in.pack(side="left", padx=2)

        btn_zoom_reset = ctk.CTkButton(zoom_frame, text="100%", width=45, height=28, font=ctk.CTkFont(size=11), fg_color="gray25", hover_color="gray35", command=self.zoom_reset)
        btn_zoom_reset.pack(side="left", padx=2)

        # Target Sample Recipient Selector
        t_lbl = ctk.CTkLabel(top_ctrl_bar, text="Preview Recipient:", font=ctk.CTkFont(size=12))
        t_lbl.grid(row=0, column=4, padx=(8, 4), pady=8, sticky="e")

        self.preview_target_menu = ctk.CTkOptionMenu(
            top_ctrl_bar,
            values=["Sample Placeholder Data"],
            width=190,
            command=lambda _: self.trigger_live_preview_update()
        )
        self.preview_target_menu.grid(row=0, column=5, padx=4, pady=8, sticky="e")

        btn_load_tpl = ctk.CTkButton(top_ctrl_bar, text="Open File", width=80, command=self.load_template_file)
        btn_load_tpl.grid(row=0, column=6, padx=3, pady=8)

        btn_save_tpl = ctk.CTkButton(top_ctrl_bar, text="Save File", width=80, fg_color="#10b981", hover_color="#059669", command=self.save_template_file)
        btn_save_tpl.grid(row=0, column=7, padx=(3, 10), pady=8)

        # Tag & Formatting Toolbar
        toolbar_frame = ctk.CTkFrame(page)
        toolbar_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        lbl_insert = ctk.CTkLabel(toolbar_frame, text="Insert Tag:", font=ctk.CTkFont(weight="bold", size=12))
        lbl_insert.pack(side="left", padx=(10, 4), pady=6)

        # Dynamic container for column tags
        self.dynamic_tags_frame = ctk.CTkFrame(toolbar_frame, fg_color="transparent")
        self.dynamic_tags_frame.pack(side="left", padx=2, pady=0)

        sep = ctk.CTkLabel(toolbar_frame, text="|", font=ctk.CTkFont(size=14), text_color="gray")
        sep.pack(side="left", padx=6)

        btn_bold = ctk.CTkButton(toolbar_frame, text="Bold <b>", width=65, height=26, font=ctk.CTkFont(size=11, weight="bold"), fg_color="gray30", hover_color="gray40", command=lambda: self.wrap_selected_html("<strong>", "</strong>"))
        btn_bold.pack(side="left", padx=2, pady=6)

        btn_italic = ctk.CTkButton(toolbar_frame, text="Italic <i>", width=65, height=26, font=ctk.CTkFont(size=11, slant="italic"), fg_color="gray30", hover_color="gray40", command=lambda: self.wrap_selected_html("<em>", "</em>"))
        btn_italic.pack(side="left", padx=2, pady=6)

        btn_u = ctk.CTkButton(toolbar_frame, text="Underline <u>", width=85, height=26, font=ctk.CTkFont(size=11), fg_color="gray30", hover_color="gray40", command=lambda: self.wrap_selected_html("<u>", "</u>"))
        btn_u.pack(side="left", padx=2, pady=6)

        btn_link = ctk.CTkButton(toolbar_frame, text="Insert Link <a>", width=95, height=26, font=ctk.CTkFont(size=11, weight="bold"), fg_color="#2563eb", hover_color="#1d4ed8", command=self.show_insert_link_modal)
        btn_link.pack(side="left", padx=2, pady=6)

        btn_p = ctk.CTkButton(toolbar_frame, text="Paragraph <p>", width=90, height=26, font=ctk.CTkFont(size=11), fg_color="gray30", hover_color="gray40", command=lambda: self.wrap_selected_html("<p>", "</p>"))
        btn_p.pack(side="left", padx=2, pady=6)

        btn_br = ctk.CTkButton(toolbar_frame, text="Break <br>", width=75, height=26, font=ctk.CTkFont(size=11), fg_color="gray30", hover_color="gray40", command=lambda: self.insert_placeholder("<br>\n"))
        btn_br.pack(side="left", padx=2, pady=6)

        self.refresh_template_tag_buttons()

        # Paned Window with Resizable Sash
        self.paned_workspace = ttk.PanedWindow(page, orient="horizontal")
        self.paned_workspace.grid(row=2, column=0, sticky="nsew")

        # LEFT PANE: VSCode Styled Code Editor
        self.editor_pane = ctk.CTkFrame(self.paned_workspace, fg_color="#1e1e1e", corner_radius=6)
        self.paned_workspace.add(self.editor_pane, weight=1)
        self.editor_pane.grid_rowconfigure(1, weight=1)
        self.editor_pane.grid_columnconfigure(0, weight=1)

        editor_header = ctk.CTkFrame(self.editor_pane, fg_color="#252526", height=28, corner_radius=0)
        editor_header.grid(row=0, column=0, sticky="ew")
        editor_header.grid_columnconfigure(0, weight=1)

        lbl_edit_title = ctk.CTkLabel(
            editor_header, 
            text="HTML Source Editor (VS Code Dark Syntax)", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#569cd6"
        )
        lbl_edit_title.grid(row=0, column=0, padx=10, pady=3, sticky="w")

        text_container = ctk.CTkFrame(self.editor_pane, fg_color="#1e1e1e", corner_radius=0)
        text_container.grid(row=1, column=0, sticky="nsew")
        text_container.grid_rowconfigure(0, weight=1)
        text_container.grid_columnconfigure(0, weight=1)

        self.editor_text = tk.Text(
            text_container,
            wrap="word",
            undo=True,
            background="#1e1e1e",
            foreground="#d4d4d4",
            insertbackground="#ffffff",
            selectbackground="#264f78",
            selectforeground="#ffffff",
            font=("Consolas", self.editor_font_size),
            padx=12,
            pady=10,
            borderwidth=0,
            highlightthickness=0
        )
        self.editor_text.grid(row=0, column=0, sticky="nsew")

        scroll_y = ttk.Scrollbar(text_container, orient="vertical", command=self.editor_text.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        self.editor_text.configure(yscrollcommand=scroll_y.set)

        self.editor_footer = ctk.CTkFrame(self.editor_pane, fg_color="#007acc", height=22, corner_radius=0)
        self.editor_footer.grid(row=2, column=0, sticky="ew")
        self.editor_footer.grid_columnconfigure(0, weight=1)

        self.lbl_cursor_pos = ctk.CTkLabel(
            self.editor_footer, 
            text="Ln 1, Col 1 | UTF-8 | HTML | Live Sync Active", 
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="white"
        )
        self.lbl_cursor_pos.grid(row=0, column=0, padx=10, pady=1, sticky="w")

        self.editor_text.bind("<KeyRelease>", self.on_editor_key_release)
        self.editor_text.bind("<ButtonRelease>", self.update_cursor_status)
        self.editor_text.bind("<Control-MouseWheel>", self.on_mouse_wheel_zoom)
        self.editor_text.bind("<Control-plus>", lambda _: self.zoom_in())
        self.editor_text.bind("<Control-equal>", lambda _: self.zoom_in())
        self.editor_text.bind("<Control-minus>", lambda _: self.zoom_out())
        self.editor_text.bind("<Control-0>", lambda _: self.zoom_reset())

        self.setup_vscode_syntax_tags()

        # RIGHT PANE: Live Visual Preview
        self.preview_pane = ctk.CTkFrame(self.paned_workspace, corner_radius=6)
        self.paned_workspace.add(self.preview_pane, weight=1)
        self.preview_pane.grid_rowconfigure(1, weight=1)
        self.preview_pane.grid_columnconfigure(0, weight=1)

        preview_header = ctk.CTkFrame(self.preview_pane, fg_color=("gray85", "gray22"), height=28, corner_radius=0)
        preview_header.grid(row=0, column=0, sticky="ew")
        preview_header.grid_columnconfigure(0, weight=1)

        lbl_prev_title = ctk.CTkLabel(
            preview_header, 
            text="Live Rendered Output (Visual Email)", 
            font=ctk.CTkFont(size=12, weight="bold"), 
            text_color="#38bdf8"
        )
        lbl_prev_title.grid(row=0, column=0, padx=10, pady=3, sticky="w")

        btn_refresh = ctk.CTkButton(
            preview_header,
            text="Refresh",
            width=65,
            height=20,
            font=ctk.CTkFont(size=10),
            fg_color="gray30",
            hover_color="gray40",
            command=self.trigger_live_preview_update
        )
        btn_refresh.grid(row=0, column=1, padx=6, pady=2, sticky="e")

        self.html_view_container = ctk.CTkFrame(self.preview_pane, fg_color="white", corner_radius=6)
        self.html_view_container.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        self.html_view_container.grid_rowconfigure(0, weight=1)
        self.html_view_container.grid_columnconfigure(0, weight=1)

        self.html_renderer_widget = None
        self.setup_html_preview_widget()

    # ==========================================
    #     PAGE 4: CAMPAIGN SCHEDULER (12H AM/PM)
    # ==========================================
    def create_schedule_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["schedule"] = page
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)

        # 1. Info Card
        info_card = ctk.CTkFrame(page)
        info_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        info_card.grid_columnconfigure(0, weight=1)

        t_lbl = ctk.CTkLabel(
            info_card, 
            text="Automated Campaign Scheduler & Dual-Timezone Queue", 
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#38bdf8"
        )
        t_lbl.grid(row=0, column=0, padx=15, pady=(10, 4), sticky="w")

        desc_text = (
            "Schedule automated outreach batches by selecting target date, 12-hour AM/PM time, and timezone.\n"
            "Both Target Timezone and Bangladesh Local Time (BST) are fully synchronized in real-time side-by-side."
        )
        lbl_d = ctk.CTkLabel(info_card, text=desc_text, font=ctk.CTkFont(size=11), justify="left", text_color="gray80")
        lbl_d.grid(row=1, column=0, padx=15, pady=(0, 8), sticky="w")

        # 2. Interactive Schedule Form (Dual Timezone Real-Time Sync)
        form_card = ctk.CTkFrame(page)
        form_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        form_card.grid_columnconfigure(0, weight=1)

        f_title = ctk.CTkLabel(form_card, text="Create New Scheduled Batch (Dual-Timezone Sync)", font=ctk.CTkFont(size=13, weight="bold"))
        f_title.pack(anchor="w", padx=15, pady=(10, 4))

        # Timezone Selector Row (Searchable Popup Button)
        tz_select_frame = ctk.CTkFrame(form_card, fg_color="transparent")
        tz_select_frame.pack(fill="x", padx=15, pady=(4, 6))

        lbl_tz = ctk.CTkLabel(tz_select_frame, text="Select Target Region / Timezone:", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_tz.pack(side="left", padx=(0, 10))

        self.sched_tz_var = ctk.StringVar(value=TIMEZONE_PRESETS[0]) # Default: US Eastern Time
        self.sched_tz_btn = ctk.CTkButton(
            tz_select_frame,
            textvariable=self.sched_tz_var,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#1e293b",
            hover_color="#334155",
            border_width=1,
            border_color="#38bdf8",
            height=34,
            anchor="w",
            command=self.open_timezone_search_modal
        )
        self.sched_tz_btn.pack(side="left", fill="x", expand=True)

        # Dual-Column Sync Card Frame
        sync_container = ctk.CTkFrame(form_card, fg_color="transparent")
        sync_container.pack(fill="x", padx=15, pady=4)
        sync_container.grid_columnconfigure(0, weight=5)
        sync_container.grid_columnconfigure(1, weight=1)
        sync_container.grid_columnconfigure(2, weight=5)

        # Left Column: Target Timezone Panel
        target_card = ctk.CTkFrame(sync_container, corner_radius=8, border_width=1, border_color="#38bdf8")
        target_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=2)

        lbl_target_header = ctk.CTkLabel(
            target_card, 
            text="🎯 Target Region Date & Time", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#38bdf8"
        )
        lbl_target_header.pack(anchor="w", padx=12, pady=(8, 4))

        # Target Date (Calendar Button)
        t_date_row = ctk.CTkFrame(target_card, fg_color="transparent")
        t_date_row.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(t_date_row, text="Date:", width=45, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")

        self.sched_target_date_var = ctk.StringVar(value=format_date_display(datetime.date.today()))
        self.sched_target_date_btn = ctk.CTkButton(
            t_date_row,
            textvariable=self.sched_target_date_var,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#1e293b",
            hover_color="#334155",
            border_width=1,
            border_color="#38bdf8",
            width=200,
            command=self.open_target_calendar
        )
        self.sched_target_date_btn.pack(side="left", fill="x", expand=True)

        # Target Time (12h AM/PM)
        t_time_row = ctk.CTkFrame(target_card, fg_color="transparent")
        t_time_row.pack(fill="x", padx=12, pady=(4, 10))
        ctk.CTkLabel(t_time_row, text="Time:", width=45, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")

        hours_12 = [f"{h:02d}" for h in range(1, 13)]
        minutes_60 = [f"{m:02d}" for m in range(0, 60, 5)]

        self.sched_target_hour_var = ctk.StringVar(value="09")
        self.sched_target_hour_menu = ctk.CTkOptionMenu(t_time_row, values=hours_12, variable=self.sched_target_hour_var, width=60, command=lambda _: self.sync_from_target())
        self.sched_target_hour_menu.pack(side="left", padx=(0, 2))

        ctk.CTkLabel(t_time_row, text=":", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=2)

        self.sched_target_min_var = ctk.StringVar(value="00")
        self.sched_target_min_menu = ctk.CTkOptionMenu(t_time_row, values=minutes_60, variable=self.sched_target_min_var, width=60, command=lambda _: self.sync_from_target())
        self.sched_target_min_menu.pack(side="left", padx=2)

        self.sched_target_ampm_var = ctk.StringVar(value="AM")
        self.sched_target_ampm_menu = ctk.CTkOptionMenu(t_time_row, values=["AM", "PM"], variable=self.sched_target_ampm_var, width=65, command=lambda _: self.sync_from_target())
        self.sched_target_ampm_menu.pack(side="left", padx=(2, 0))

        # Middle Column: Sync Indicator
        mid_card = ctk.CTkFrame(sync_container, fg_color="transparent")
        mid_card.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        lbl_sync_icon = ctk.CTkLabel(
            mid_card, 
            text="⇄\nAuto-Sync", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#10b981"
        )
        lbl_sync_icon.pack(expand=True)

        # Right Column: Bangladesh Time Panel
        bd_card = ctk.CTkFrame(sync_container, corner_radius=8, border_width=1, border_color="#10b981")
        bd_card.grid(row=0, column=2, sticky="nsew", padx=(6, 0), pady=2)

        lbl_bd_header = ctk.CTkLabel(
            bd_card, 
            text="🇧🇩 Bangladesh Time (BST - UTC+6)", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#10b981"
        )
        lbl_bd_header.pack(anchor="w", padx=12, pady=(8, 4))

        # BD Date (Calendar Button)
        bd_date_row = ctk.CTkFrame(bd_card, fg_color="transparent")
        bd_date_row.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(bd_date_row, text="Date:", width=45, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")

        self.sched_bd_date_var = ctk.StringVar(value=format_date_display(datetime.date.today()))
        self.sched_bd_date_btn = ctk.CTkButton(
            bd_date_row,
            textvariable=self.sched_bd_date_var,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#1e293b",
            hover_color="#334155",
            border_width=1,
            border_color="#10b981",
            width=200,
            command=self.open_bd_calendar
        )
        self.sched_bd_date_btn.pack(side="left", fill="x", expand=True)

        # BD Time (12h AM/PM)
        bd_time_row = ctk.CTkFrame(bd_card, fg_color="transparent")
        bd_time_row.pack(fill="x", padx=12, pady=(4, 10))
        ctk.CTkLabel(bd_time_row, text="Time:", width=45, anchor="w", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")

        self.sched_bd_hour_var = ctk.StringVar(value="07")
        self.sched_bd_hour_menu = ctk.CTkOptionMenu(bd_time_row, values=hours_12, variable=self.sched_bd_hour_var, width=60, command=lambda _: self.sync_from_bd())
        self.sched_bd_hour_menu.pack(side="left", padx=(0, 2))

        ctk.CTkLabel(bd_time_row, text=":", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=2)

        self.sched_bd_min_var = ctk.StringVar(value="00")
        self.sched_bd_min_menu = ctk.CTkOptionMenu(bd_time_row, values=minutes_60, variable=self.sched_bd_min_var, width=60, command=lambda _: self.sync_from_bd())
        self.sched_bd_min_menu.pack(side="left", padx=2)

        self.sched_bd_ampm_var = ctk.StringVar(value="PM")
        self.sched_bd_ampm_menu = ctk.CTkOptionMenu(bd_time_row, values=["AM", "PM"], variable=self.sched_bd_ampm_var, width=65, command=lambda _: self.sync_from_bd())
        self.sched_bd_ampm_menu.pack(side="left", padx=(2, 0))

        # Row 3: Action Buttons & Live Preview Banner
        action_row = ctk.CTkFrame(form_card, fg_color="transparent")
        action_row.pack(fill="x", padx=15, pady=(6, 12))

        # Row Range Filter
        rows_frame = ctk.CTkFrame(action_row, fg_color="transparent")
        rows_frame.pack(side="left")

        ctk.CTkLabel(rows_frame, text="Recipient Rows:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 6))
        self.sched_row_from = ctk.CTkEntry(rows_frame, width=50, placeholder_text="1")
        self.sched_row_from.insert(0, "1")
        self.sched_row_from.pack(side="left", padx=2)

        ctk.CTkLabel(rows_frame, text="to").pack(side="left", padx=4)
        self.sched_row_to = ctk.CTkEntry(rows_frame, width=55, placeholder_text="End")
        self.sched_row_to.pack(side="left", padx=2)

        # Live Sync Preview Banner in middle
        self.lbl_calculated_time = ctk.CTkLabel(
            action_row, 
            text="Preview: Synchronizing...", 
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_calculated_time.pack(side="left", padx=15, expand=True)

        # Add Batch Button on right
        btn_add_schedule = ctk.CTkButton(
            action_row, 
            text="⏰ Add to Scheduled List", 
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#10b981", 
            hover_color="#059669",
            width=185,
            command=self.add_batch_to_scheduled_queue
        )
        btn_add_schedule.pack(side="right", padx=2)

        # Initial sync from Target
        self.sync_from_target()

        # 3. Scheduled Batches Queue Table
        queue_container = ctk.CTkFrame(page)
        queue_container.grid(row=2, column=0, sticky="nsew")
        queue_container.grid_rowconfigure(1, weight=1)
        queue_container.grid_columnconfigure(0, weight=1)

        q_header = ctk.CTkFrame(queue_container, fg_color="transparent")
        q_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        q_header.grid_columnconfigure(0, weight=1)

        lbl_q_title = ctk.CTkLabel(q_header, text="Active Scheduled Queue (Auto-Execution)", font=ctk.CTkFont(size=13, weight="bold"))
        lbl_q_title.grid(row=0, column=0, sticky="w")

        btn_del_batch = ctk.CTkButton(
            q_header, 
            text="Remove Selected", 
            height=26, 
            fg_color="#dc2626", 
            hover_color="#b91c1c",
            command=self.delete_selected_batch
        )
        btn_del_batch.grid(row=0, column=1, padx=4)

        btn_clear_all_q = ctk.CTkButton(
            q_header, 
            text="Clear Queue", 
            height=26, 
            fg_color="gray30", 
            hover_color="gray40",
            command=self.clear_all_scheduled_batches
        )
        btn_clear_all_q.grid(row=0, column=2, padx=4)

        self.queue_tree = ttk.Treeview(
            queue_container, 
            style="Custom.Treeview", 
            columns=("ID", "Zone", "Target_Time", "BD_Time", "Countdown", "Rows", "Status"), 
            show="headings"
        )
        self.queue_tree.heading("ID", text="Batch ID")
        self.queue_tree.heading("Zone", text="Target Timezone")
        self.queue_tree.heading("Target_Time", text="Target Region Time")
        self.queue_tree.heading("BD_Time", text="BD Execution Time (BST)")
        self.queue_tree.heading("Countdown", text="Countdown")
        self.queue_tree.heading("Rows", text="Target Rows")
        self.queue_tree.heading("Status", text="Status")

        self.queue_tree.column("ID", width=65, anchor="center")
        self.queue_tree.column("Zone", width=190)
        self.queue_tree.column("Target_Time", width=180, anchor="center")
        self.queue_tree.column("BD_Time", width=180, anchor="center")
        self.queue_tree.column("Countdown", width=120, anchor="center")
        self.queue_tree.column("Rows", width=90, anchor="center")
        self.queue_tree.column("Status", width=100, anchor="center")

        q_scroll = ttk.Scrollbar(queue_container, orient="vertical", command=self.queue_tree.yview)
        self.queue_tree.configure(yscrollcommand=q_scroll.set)

        self.queue_tree.grid(row=1, column=0, sticky="nsew", padx=(10, 0), pady=(0, 10))
        q_scroll.grid(row=1, column=1, sticky="ns", padx=(0, 10), pady=(0, 10))

    def open_target_calendar(self):
        target_tz = get_timezone_obj(self.sched_tz_var.get())
        now_in_target = datetime.datetime.now(target_tz).date()
        
        CalendarPickerModal(
            parent=self,
            initial_date=self.sched_target_date_var.get(),
            min_date=now_in_target,
            title="🎯 Select Target Region Date",
            callback=self.on_target_calendar_selected
        )

    def on_target_calendar_selected(self, date_str):
        self.sched_target_date_var.set(format_date_display(date_str))
        self.sync_from_target()

    def open_bd_calendar(self):
        now_in_bd = datetime.datetime.now(BD_TIMEZONE).date()
        
        CalendarPickerModal(
            parent=self,
            initial_date=self.sched_bd_date_var.get(),
            min_date=now_in_bd,
            title="🇧🇩 Select Bangladesh Date",
            callback=self.on_bd_calendar_selected
        )

    def on_bd_calendar_selected(self, date_str):
        self.sched_bd_date_var.set(format_date_display(date_str))
        self.sync_from_bd()

    def parse_picker_datetime(self, date_var, hour_var, min_var, ampm_var, tz_obj):
        """Parses date and 12h dropdowns into an aware datetime object in the specified timezone"""
        date_raw = date_var.get()
        m = re.search(r'(\d{4}-\d{2}-\d{2})', date_raw)
        if not m:
            return None
        date_str = m.group(1)

        try:
            hour_12 = int(hour_var.get())
            minute = int(min_var.get())
            ampm = ampm_var.get().upper()

            if ampm == "PM" and hour_12 < 12:
                hour_24 = hour_12 + 12
            elif ampm == "AM" and hour_12 == 12:
                hour_24 = 0
            else:
                hour_24 = hour_12

            dt_naive = datetime.datetime.strptime(f"{date_str} {hour_24:02d}:{minute:02d}", "%Y-%m-%d %H:%M")
            return dt_naive.replace(tzinfo=tz_obj)
        except Exception as e:
            return None

    def open_timezone_search_modal(self):
        TimezoneSearchModal(
            parent=self,
            timezone_list=TIMEZONE_PRESETS,
            selected_tz=self.sched_tz_var.get(),
            callback=self.on_timezone_selected
        )

    def on_timezone_selected(self, new_tz):
        self.sched_tz_var.set(new_tz)
        self.sync_from_bd()

    def sync_from_target(self):
        if getattr(self, "_syncing_timezones", False):
            return
        self._syncing_timezones = True
        try:
            target_tz = get_timezone_obj(self.sched_tz_var.get())
            dt_target = self.parse_picker_datetime(
                self.sched_target_date_var,
                self.sched_target_hour_var,
                self.sched_target_min_var,
                self.sched_target_ampm_var,
                target_tz
            )
            if dt_target:
                dt_bd = dt_target.astimezone(BD_TIMEZONE)
                
                self.sched_bd_date_var.set(format_date_display(dt_bd.strftime("%Y-%m-%d")))
                
                bd_hour_12 = dt_bd.hour % 12 or 12
                self.sched_bd_hour_var.set(f"{bd_hour_12:02d}")
                
                rounded_min = (dt_bd.minute // 5) * 5
                self.sched_bd_min_var.set(f"{rounded_min:02d}")
                
                self.sched_bd_ampm_var.set("PM" if dt_bd.hour >= 12 else "AM")
                
                self.update_live_sync_banner(dt_target, dt_bd)
        except Exception as e:
            print(f"Sync from target error: {e}")
        finally:
            self._syncing_timezones = False

    def sync_from_bd(self):
        if getattr(self, "_syncing_timezones", False):
            return
        self._syncing_timezones = True
        try:
            dt_bd = self.parse_picker_datetime(
                self.sched_bd_date_var,
                self.sched_bd_hour_var,
                self.sched_bd_min_var,
                self.sched_bd_ampm_var,
                BD_TIMEZONE
            )
            if dt_bd:
                target_tz = get_timezone_obj(self.sched_tz_var.get())
                dt_target = dt_bd.astimezone(target_tz)
                
                self.sched_target_date_var.set(format_date_display(dt_target.strftime("%Y-%m-%d")))
                
                target_hour_12 = dt_target.hour % 12 or 12
                self.sched_target_hour_var.set(f"{target_hour_12:02d}")
                
                rounded_min = (dt_target.minute // 5) * 5
                self.sched_target_min_var.set(f"{rounded_min:02d}")
                
                self.sched_target_ampm_var.set("PM" if dt_target.hour >= 12 else "AM")
                
                self.update_live_sync_banner(dt_target, dt_bd)
        except Exception as e:
            print(f"Sync from BD error: {e}")
        finally:
            self._syncing_timezones = False

    def update_live_sync_banner(self, dt_target, dt_bd):
        try:
            t_str = dt_target.strftime("%a, %b %d at %I:%M %p")
            bd_str = dt_bd.strftime("%a, %b %d at %I:%M %p BST")
            
            now_bd = datetime.datetime.now(BD_TIMEZONE)
            diff_sec = (dt_bd - now_bd).total_seconds()
            if diff_sec > 0:
                h = int(diff_sec // 3600)
                m = int((diff_sec % 3600) // 60)
                cd = f"⏰ Triggers in: {h}h {m}m"
                color = "#38bdf8"
            else:
                cd = "⚠️ Invalid: Selected time is in the past!"
                color = "#f87171"
                
            self.lbl_calculated_time.configure(
                text=f"🎯 Target: {t_str}  ⇄  🇧🇩 BD: {bd_str}  |  {cd}",
                text_color=color
            )
        except Exception as e:
            pass

    def get_selected_both_datetimes(self):
        target_tz = get_timezone_obj(self.sched_tz_var.get())
        dt_target = self.parse_picker_datetime(
            self.sched_target_date_var,
            self.sched_target_hour_var,
            self.sched_target_min_var,
            self.sched_target_ampm_var,
            target_tz
        )
        dt_bd = self.parse_picker_datetime(
            self.sched_bd_date_var,
            self.sched_bd_hour_var,
            self.sched_bd_min_var,
            self.sched_bd_ampm_var,
            BD_TIMEZONE
        )
        return dt_target, dt_bd

    def add_batch_to_scheduled_queue(self):
        if self.recipients_df is None or self.recipients_df.empty:
            messagebox.showwarning("No Data", "Please load a CSV or Excel recipient list first.")
            return

        dt_target, dt_bd = self.get_selected_both_datetimes()
        if not dt_bd or not dt_target:
            messagebox.showerror("Invalid Date/Time", "Please select a valid date and time.")
            return

        now_bd = datetime.datetime.now(BD_TIMEZONE)
        if dt_bd <= now_bd:
            messagebox.showerror(
                "Invalid Time: Past DateTime",
                f"The selected scheduled execution time is in the past:\n\n"
                f"Selected Execution (BD Time): {dt_bd.strftime('%b %d, %Y %I:%M %p BST')}\n"
                f"Current Bangladesh Time: {now_bd.strftime('%b %d, %Y %I:%M %p BST')}\n\n"
                f"Please choose a future date or time."
            )
            return

        total_rows = len(self.recipients_df)
        r_from = int(self.sched_row_from.get().strip() or 1)
        r_to_str = self.sched_row_to.get().strip()
        r_to = int(r_to_str) if r_to_str.isdigit() else total_rows
        r_to = min(r_to, total_rows)

        batch_id = f"B{len(self.scheduled_batches)+1:03d}"
        target_tz_name = self.sched_tz_var.get()
        target_disp = dt_target.strftime("%b %d, %Y %I:%M %p")
        bd_disp = dt_bd.strftime("%b %d, %Y %I:%M %p BST")
        target_iso = dt_bd.strftime("%Y-%m-%d %H:%M")
        target_timestamp = dt_bd.timestamp()

        new_batch = {
            "id": batch_id,
            "timezone": target_tz_name,
            "target_region_time": target_disp,
            "bd_execution_time": bd_disp,
            "target_time": target_iso,
            "target_timestamp": target_timestamp,
            "display_time": f"{target_disp} ({bd_disp})",
            "row_from": r_from,
            "row_to": r_to,
            "status": "Pending",
            "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        self.scheduled_batches.append(new_batch)
        self.save_scheduled_queue()
        self.refresh_scheduled_batches_table()
        self.log_message(f"[+] Scheduled batch {batch_id} added: Target {target_disp} (BD Exec: {bd_disp}) [Rows {r_from}-{r_to}]", "info")
        messagebox.showinfo("Batch Queued", f"Batch {batch_id} successfully added to scheduled list!\n\n🎯 Target Region Time:\n{target_disp} ({target_tz_name})\n\n🇧🇩 Bangladesh Execution Time:\n{bd_disp}\n\nTarget Rows: {r_from} to {r_to}")

    def refresh_scheduled_batches_table(self):
        self.queue_tree.delete(*self.queue_tree.get_children())
        now_ts = time.time()

        for b in self.scheduled_batches:
            try:
                ts = b.get("target_timestamp")
                if not ts:
                    target_dt = datetime.datetime.strptime(b["target_time"], "%Y-%m-%d %H:%M")
                    ts = target_dt.replace(tzinfo=BD_TIMEZONE).timestamp()
                
                diff = ts - now_ts
                if diff > 0 and b.get("status") == "Pending":
                    hours = int(diff // 3600)
                    mins = int((diff % 3600) // 60)
                    secs = int(diff % 60)
                    cd_str = f"in {hours}h {mins}m {secs}s"
                elif b.get("status") == "Completed":
                    cd_str = "Completed"
                elif b.get("status") == "Running":
                    cd_str = "Running Now"
                elif b.get("status") == "Failed":
                    cd_str = "Failed"
                elif b.get("status") == "Stopped":
                    cd_str = "Stopped"
                else:
                    cd_str = "Ready / Due"
            except Exception:
                cd_str = "—"

            rows_str = f"{b.get('row_from', 1)}-{b.get('row_to', 'End')}"
            target_disp = b.get("target_region_time", b.get("display_time", b.get("target_time", "")))
            bd_disp = b.get("bd_execution_time", b.get("display_time", b.get("target_time", "")))
            zone_disp = b.get("timezone", "Default")
            if len(zone_disp) > 32:
                zone_disp = zone_disp.split("(")[0].strip()

            self.queue_tree.insert("", tk.END, values=(
                b["id"], 
                zone_disp, 
                target_disp,
                bd_disp,
                cd_str, 
                rows_str, 
                b.get("status", "Pending")
            ))

    def delete_selected_batch(self):
        selected = self.queue_tree.selection()
        if not selected:
            messagebox.showwarning("Select Batch", "Please select a batch from the table to remove.")
            return

        item = self.queue_tree.item(selected[0])
        batch_id = item["values"][0]

        self.scheduled_batches = [b for b in self.scheduled_batches if b["id"] != batch_id]
        self.save_scheduled_queue()
        self.refresh_scheduled_batches_table()
        self.log_message(f"[!] Removed scheduled batch {batch_id}", "info")

    def clear_all_scheduled_batches(self):
        if messagebox.askyesno("Clear Queue", "Are you sure you want to clear all queued batches?"):
            self.scheduled_batches = []
            self.save_scheduled_queue()
            self.refresh_scheduled_batches_table()
            self.log_message("[!] Cleared all scheduled batches.", "info")

    # ==========================================
    #     PAGE 5: ANTI-SPAM & DELIVERY SETTINGS
    # ==========================================
    def create_antispam_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["antispam"] = page
        page.grid_columnconfigure(0, weight=1)

        throttle_card = ctk.CTkFrame(page)
        throttle_card.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        throttle_card.grid_columnconfigure(1, weight=1)

        t_title = ctk.CTkLabel(
            throttle_card, 
            text="Anti-Spam Rate Limiting & Delivery Controls", 
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#10b981"
        )
        t_title.grid(row=0, column=0, columnspan=2, padx=20, pady=(15, 10), sticky="w")

        lbl_delay = ctk.CTkLabel(throttle_card, text="Base Delay between emails (seconds):", font=ctk.CTkFont(weight="bold"))
        lbl_delay.grid(row=1, column=0, padx=20, pady=8, sticky="w")

        self.delay_spinbox = ctk.CTkEntry(throttle_card, width=110)
        self.delay_spinbox.insert(0, str(self.config_data.get("delay_seconds", 12)))
        self.delay_spinbox.grid(row=1, column=1, padx=20, pady=8, sticky="w")

        lbl_jitter = ctk.CTkLabel(throttle_card, text="Random Delay Jitter (± seconds):", font=ctk.CTkFont(weight="bold"))
        lbl_jitter.grid(row=2, column=0, padx=20, pady=8, sticky="w")

        self.jitter_spinbox = ctk.CTkEntry(throttle_card, width=110)
        self.jitter_spinbox.insert(0, str(self.config_data.get("random_jitter", 4)))
        self.jitter_spinbox.grid(row=2, column=1, padx=20, pady=8, sticky="w")

        lbl_batch_count = ctk.CTkLabel(throttle_card, text="Cooldown Pause after every X emails:", font=ctk.CTkFont(weight="bold"))
        lbl_batch_count.grid(row=3, column=0, padx=20, pady=8, sticky="w")

        self.batch_count_spinbox = ctk.CTkEntry(throttle_card, width=110)
        self.batch_count_spinbox.insert(0, str(self.config_data.get("batch_pause_count", 20)))
        self.batch_count_spinbox.grid(row=3, column=1, padx=20, pady=8, sticky="w")

        lbl_batch_pause = ctk.CTkLabel(throttle_card, text="Cooldown Pause Duration (minutes):", font=ctk.CTkFont(weight="bold"))
        lbl_batch_pause.grid(row=4, column=0, padx=20, pady=8, sticky="w")

        self.batch_pause_spinbox = ctk.CTkEntry(throttle_card, width=110)
        self.batch_pause_spinbox.insert(0, str(self.config_data.get("batch_pause_minutes", 5)))
        self.batch_pause_spinbox.grid(row=4, column=1, padx=20, pady=8, sticky="w")

        lbl_daily = ctk.CTkLabel(throttle_card, text="Daily Sending Quota (emails/day):", font=ctk.CTkFont(weight="bold"))
        lbl_daily.grid(row=5, column=0, padx=20, pady=8, sticky="w")

        self.daily_limit_spinbox = ctk.CTkEntry(throttle_card, width=110)
        self.daily_limit_spinbox.insert(0, str(self.config_data.get("daily_limit", 150)))
        self.daily_limit_spinbox.grid(row=5, column=1, padx=20, pady=8, sticky="w")

        btn_save_timer = ctk.CTkButton(
            throttle_card, 
            text="Save Delivery Settings", 
            fg_color="#10b981", 
            hover_color="#059669",
            command=self.save_settings_action
        )
        btn_save_timer.grid(row=6, column=0, columnspan=2, padx=20, pady=(15, 20), sticky="w")

        guide_box = ctk.CTkFrame(page)
        guide_box.grid(row=1, column=0, sticky="ew")

        g_title = ctk.CTkLabel(guide_box, text="Deliverability Recommendations:", font=ctk.CTkFont(weight="bold"))
        g_title.pack(anchor="w", padx=20, pady=(12, 4))

        g_text = (
            "• Maintaining a 10s-15s interval with randomized jitter avoids automated spam triggers.\n"
            "• Taking a 5-minute cooldown break every 20 emails ensures consistent SMTP server trust.\n"
            "• Recommended daily volume: 50 to 150 emails/day for free Gmail accounts."
        )
        lbl_gt = ctk.CTkLabel(guide_box, text=g_text, font=ctk.CTkFont(size=11), justify="left", text_color="gray")
        lbl_gt.pack(anchor="w", padx=20, pady=(0, 15))

    # ==========================================
    #           PAGE 6: SMTP SETTINGS
    # ==========================================
    def create_settings_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["settings"] = page
        page.grid_columnconfigure(0, weight=1)

        smtp_box = ctk.CTkFrame(page)
        smtp_box.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        smtp_box.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(smtp_box, text="SMTP Mail Server Configuration", font=ctk.CTkFont(size=16, weight="bold"))
        title.grid(row=0, column=0, columnspan=2, padx=20, pady=(15, 10), sticky="w")

        p_lbl = ctk.CTkLabel(smtp_box, text="Email Provider Preset:")
        p_lbl.grid(row=1, column=0, padx=20, pady=6, sticky="w")

        self.smtp_provider_var = ctk.StringVar(value=self.config_data.get("smtp_provider", "Gmail"))
        self.provider_menu = ctk.CTkOptionMenu(
            smtp_box, 
            values=list(SMTP_PRESETS.keys()),
            variable=self.smtp_provider_var,
            command=self.on_provider_change
        )
        self.provider_menu.grid(row=1, column=1, padx=20, pady=6, sticky="w")

        e_lbl = ctk.CTkLabel(smtp_box, text="Sender Email Address:")
        e_lbl.grid(row=2, column=0, padx=20, pady=6, sticky="w")

        self.smtp_email_entry = ctk.CTkEntry(smtp_box, placeholder_text="yourname@gmail.com")
        self.smtp_email_entry.insert(0, self.config_data.get("email", ""))
        self.smtp_email_entry.grid(row=2, column=1, padx=20, pady=6, sticky="ew")

        pass_lbl = ctk.CTkLabel(smtp_box, text="App Password / Token:")
        pass_lbl.grid(row=3, column=0, padx=20, pady=6, sticky="w")

        pass_frame = ctk.CTkFrame(smtp_box, fg_color="transparent")
        pass_frame.grid(row=3, column=1, padx=20, pady=6, sticky="ew")
        pass_frame.grid_columnconfigure(0, weight=1)

        self.smtp_pass_entry = ctk.CTkEntry(pass_frame, placeholder_text="16-character App Password (e.g. abcd efgh ijkl mnop)", show="•")
        self.smtp_pass_entry.insert(0, self.config_data.get("password", ""))
        self.smtp_pass_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.btn_toggle_pass = ctk.CTkButton(
            pass_frame, 
            text="View", 
            width=50, 
            command=self.toggle_pass_visibility
        )
        self.btn_toggle_pass.grid(row=0, column=1)

        name_lbl = ctk.CTkLabel(smtp_box, text="Sender Display Name (Optional):")
        name_lbl.grid(row=4, column=0, padx=20, pady=6, sticky="w")

        self.sender_name_entry = ctk.CTkEntry(smtp_box, placeholder_text="e.g. S. S. M Jahir Jahan Khan Miru")
        self.sender_name_entry.insert(0, self.config_data.get("sender_name", ""))
        self.sender_name_entry.grid(row=4, column=1, padx=20, pady=6, sticky="ew")

        host_lbl = ctk.CTkLabel(smtp_box, text="SMTP Host & Port:")
        host_lbl.grid(row=5, column=0, padx=20, pady=6, sticky="w")

        hp_frame = ctk.CTkFrame(smtp_box, fg_color="transparent")
        hp_frame.grid(row=5, column=1, padx=20, pady=6, sticky="w")

        self.smtp_host_entry = ctk.CTkEntry(hp_frame, width=180)
        self.smtp_host_entry.insert(0, self.config_data.get("smtp_host", "smtp.gmail.com"))
        self.smtp_host_entry.grid(row=0, column=0, padx=(0, 8))

        self.smtp_port_entry = ctk.CTkEntry(hp_frame, width=70)
        self.smtp_port_entry.insert(0, str(self.config_data.get("smtp_port", 587)))
        self.smtp_port_entry.grid(row=0, column=1, padx=4)

        self.smtp_tls_var = ctk.BooleanVar(value=self.config_data.get("smtp_tls", True))
        self.smtp_tls_cb = ctk.CTkCheckBox(hp_frame, text="Use STARTTLS", variable=self.smtp_tls_var)
        self.smtp_tls_cb.grid(row=0, column=2, padx=12)

        btn_frame = ctk.CTkFrame(smtp_box, fg_color="transparent")
        btn_frame.grid(row=6, column=0, columnspan=2, padx=20, pady=20, sticky="ew")

        btn_test_conn = ctk.CTkButton(
            btn_frame, 
            text="Test SMTP Connection", 
            fg_color="#0284c7",
            hover_color="#0369a1",
            command=self.test_smtp_connection
        )
        btn_test_conn.grid(row=0, column=0, padx=(0, 10))

        btn_save_settings = ctk.CTkButton(
            btn_frame, 
            text="Save Credentials", 
            fg_color="#10b981",
            hover_color="#059669",
            command=self.save_settings_action
        )
        btn_save_settings.grid(row=0, column=1, padx=4)

        help_box = ctk.CTkFrame(page)
        help_box.grid(row=1, column=0, sticky="ew")
        
        h_title = ctk.CTkLabel(help_box, text="Google App Password Setup Guide:", font=ctk.CTkFont(weight="bold"))
        h_title.grid(row=0, column=0, padx=20, pady=(12, 4), sticky="w")

        h_text = ctk.CTkLabel(
            help_box, 
            text="1. Visit Google Account > Security > 2-Step Verification.\n"
                 "2. Navigate to 'App passwords' at the bottom.\n"
                 "3. Generate a password named 'AutoMailer' and paste the 16-character code here.\n"
                 "4. Note: Do not use your standard Google account login password.",
            justify="left",
            text_color="gray"
        )
        h_text.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")

    def toggle_pass_visibility(self):
        if self.smtp_pass_entry.cget("show") == "•":
            self.smtp_pass_entry.configure(show="")
            self.btn_toggle_pass.configure(text="Hide")
        else:
            self.smtp_pass_entry.configure(show="•")
            self.btn_toggle_pass.configure(text="View")

    def on_provider_change(self, choice):
        if choice in SMTP_PRESETS and choice != "Custom SMTP":
            preset = SMTP_PRESETS[choice]
            self.smtp_host_entry.delete(0, tk.END)
            self.smtp_host_entry.insert(0, preset["host"])
            self.smtp_port_entry.delete(0, tk.END)
            self.smtp_port_entry.insert(0, str(preset["port"]))
            self.smtp_tls_var.set(preset["tls"])

    def test_smtp_connection(self):
        email = self.smtp_email_entry.get().strip()
        pwd = self.smtp_pass_entry.get().strip()
        host = self.smtp_host_entry.get().strip()
        port = int(self.smtp_port_entry.get().strip() or 587)
        use_tls = self.smtp_tls_var.get()

        if not email or not pwd:
            messagebox.showwarning("Missing Info", "Please enter both Email and App Password.")
            return

        def run_test():
            try:
                self.log_message(f"Connecting to {host}:{port}...", "info")
                server = smtplib.SMTP(host, port, timeout=15)
                if use_tls:
                    server.starttls()
                server.login(email, pwd)
                server.quit()
                self.log_message("[+] SMTP Authentication Successful!", "success")
                messagebox.showinfo("Success", "SMTP Connection & Authentication Successful!")
            except Exception as e:
                self.log_message(f"[-] SMTP Connection Failed: {e}", "error")
                messagebox.showerror("Connection Error", f"Failed to authenticate:\n{e}")

        threading.Thread(target=run_test, daemon=True).start()

    def save_settings_action(self):
        self.save_config()
        messagebox.showinfo("Saved", "Configuration and credentials saved successfully!")

    # ==========================================
    #           PAGE 7: LOGS & REPORTS
    # ==========================================
    def create_logs_page(self):
        page = ctk.CTkFrame(self.pages_container, fg_color="transparent")
        self.pages["logs"] = page
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)

        top_bar = ctk.CTkFrame(page)
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        lbl = ctk.CTkLabel(top_bar, text="Activity & Campaign Log Console", font=ctk.CTkFont(weight="bold"))
        lbl.grid(row=0, column=0, padx=12, pady=10, sticky="w")

        btn_export = ctk.CTkButton(
            top_bar, 
            text="Export Report to CSV", 
            fg_color="#0284c7",
            command=self.export_report_csv
        )
        btn_export.grid(row=0, column=1, padx=6, pady=10)

        btn_clear = ctk.CTkButton(
            top_bar, 
            text="Clear Logs", 
            fg_color="gray30",
            hover_color="gray40",
            command=self.clear_logs
        )
        btn_clear.grid(row=0, column=2, padx=6, pady=10)

        self.logs_textbox = ctk.CTkTextbox(
            page, 
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word"
        )
        self.logs_textbox.grid(row=1, column=0, sticky="nsew")

    def log_message(self, message, msg_type="info"):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {message}\n"
        
        def _update():
            if hasattr(self, 'logs_textbox') and self.logs_textbox.winfo_exists():
                self.logs_textbox.insert(tk.END, formatted)
                self.logs_textbox.see(tk.END)
            if hasattr(self, 'quick_log_text') and self.quick_log_text.winfo_exists():
                self.quick_log_text.insert(tk.END, formatted)
                self.quick_log_text.see(tk.END)

        self.after(0, _update)

    def clear_logs(self):
        self.logs_textbox.delete("1.0", tk.END)
        self.quick_log_text.delete("1.0", tk.END)

    def export_report_csv(self):
        if not self.campaign_log_records:
            messagebox.showinfo("No Data", "No campaign results to export yet.")
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv", 
            filetypes=[("CSV File", "*.csv")],
            initialfile=f"Campaign_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if save_path:
            try:
                rep_df = pd.DataFrame(self.campaign_log_records)
                rep_df.to_csv(save_path, index=False)
                messagebox.showinfo("Exported", f"Report successfully saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Export Failed", str(e))

    # ==========================================
    #     VS CODE SYNTAX HIGHLIGHTING ENGINE
    # ==========================================
    def setup_vscode_syntax_tags(self):
        f_size = self.editor_font_size
        self.editor_text.tag_configure("tag_name", foreground="#569CD6")
        self.editor_text.tag_configure("tag_bracket", foreground="#808080")
        self.editor_text.tag_configure("attr_name", foreground="#9CDCFE")
        self.editor_text.tag_configure("attr_val", foreground="#CE9178")
        self.editor_text.tag_configure("comment", foreground="#6A9955", font=("Consolas", f_size, "italic"))
        self.editor_text.tag_configure("variable", foreground="#FFD700", background="#2d2b14", font=("Consolas", f_size, "bold"))
        self.editor_text.tag_configure("entity", foreground="#D7BA7D")

    def apply_syntax_highlighting(self):
        if not hasattr(self, 'editor_text') or not self.editor_text.winfo_exists():
            return

        for tag in ["tag_name", "tag_bracket", "attr_name", "attr_val", "comment", "variable", "entity"]:
            self.editor_text.tag_remove(tag, "1.0", "end")

        content = self.editor_text.get("1.0", "end")

        for match in re.finditer(r'<!--.*?-->', content, re.DOTALL):
            s = f"1.0 + {match.start()} chars"
            e = f"1.0 + {match.end()} chars"
            self.editor_text.tag_add("comment", s, e)

        for match in re.finditer(r'\{[A-Za-z0-9_ ]+\}', content):
            s = f"1.0 + {match.start()} chars"
            e = f"1.0 + {match.end()} chars"
            self.editor_text.tag_add("variable", s, e)

        for match in re.finditer(r'&[a-zA-Z0-9#]+;', content):
            s = f"1.0 + {match.start()} chars"
            e = f"1.0 + {match.end()} chars"
            self.editor_text.tag_add("entity", s, e)

        for match in re.finditer(r'<(/?[a-zA-Z0-9\-!]+)([^>]*)>', content):
            tag_start = match.start()
            tag_name_str = match.group(1)
            
            s = f"1.0 + {tag_start} chars"
            e = f"1.0 + {tag_start + 1 + len(tag_name_str)} chars"
            self.editor_text.tag_add("tag_name", s, e)

            close_s = f"1.0 + {match.end() - 1} chars"
            close_e = f"1.0 + {match.end()} chars"
            self.editor_text.tag_add("tag_bracket", close_s, close_e)

            attrs_str = match.group(2)
            attr_offset = tag_start + 1 + len(tag_name_str)

            for val_match in re.finditer(r'=\s*("[^"]*"|\'[^\']*\')', attrs_str):
                vs = f"1.0 + {attr_offset + val_match.start(1)} chars"
                ve = f"1.0 + {attr_offset + val_match.end(1)} chars"
                self.editor_text.tag_add("attr_val", vs, ve)

            for name_match in re.finditer(r'([a-zA-Z\-:]+)\s*=', attrs_str):
                ns = f"1.0 + {attr_offset + name_match.start(1)} chars"
                ne = f"1.0 + {attr_offset + name_match.end(1)} chars"
                self.editor_text.tag_add("attr_name", ns, ne)

    def set_zoom_level(self, new_size):
        self.editor_font_size = max(8, min(36, new_size))
        self.editor_text.configure(font=("Consolas", self.editor_font_size))
        self.setup_vscode_syntax_tags()
        self.lbl_zoom.configure(text=f"{int((self.editor_font_size/12)*100)}%")
        self.apply_syntax_highlighting()
        self.save_config()

    def zoom_in(self):
        self.set_zoom_level(self.editor_font_size + 1)

    def zoom_out(self):
        self.set_zoom_level(self.editor_font_size - 1)

    def zoom_reset(self):
        self.set_zoom_level(12)

    def on_mouse_wheel_zoom(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()
        return "break"

    def update_cursor_status(self, event=None):
        try:
            line, col = self.editor_text.index(tk.INSERT).split(".")
            chars = len(self.editor_text.get("1.0", "end-1c"))
            self.lbl_cursor_pos.configure(
                text=f"Ln {line}, Col {int(col)+1} | {chars} chars | UTF-8 | HTML | Live Sync Active"
            )
        except Exception:
            pass

    def on_view_mode_changed(self, choice):
        for pane in list(self.paned_workspace.panes()):
            self.paned_workspace.forget(pane)

        if "Split" in choice:
            self.paned_workspace.add(self.editor_pane, weight=1)
            self.paned_workspace.add(self.preview_pane, weight=1)
            self.trigger_live_preview_update()
        elif "Editor" in choice:
            self.paned_workspace.add(self.editor_pane, weight=1)
        elif "Preview" in choice:
            self.paned_workspace.add(self.preview_pane, weight=1)
            self.trigger_live_preview_update()

    def on_editor_key_release(self, event=None):
        self.update_cursor_status()
        if self._live_preview_after_id:
            self.after_cancel(self._live_preview_after_id)
        self._live_preview_after_id = self.after(150, self.render_live_html_preview)

        if self._syntax_highlight_after_id:
            self.after_cancel(self._syntax_highlight_after_id)
        self._syntax_highlight_after_id = self.after(100, self.apply_syntax_highlighting)

    def trigger_live_preview_update(self):
        self.render_live_html_preview()

    def setup_html_preview_widget(self):
        for child in self.html_view_container.winfo_children():
            child.destroy()

        if HAS_TKINTERWEB:
            self.html_renderer_widget = HtmlFrame(
                self.html_view_container, 
                horizontal_scrollbar="auto",
                vertical_scrollbar="auto"
            )
            self.html_renderer_widget.pack(fill="both", expand=True)
        elif HAS_TKHTMLVIEW:
            self.html_renderer_widget = tkhtmlview.HTMLText(
                self.html_view_container,
                background="white",
                foreground="black",
                font=("Arial", 11)
            )
            self.html_renderer_widget.pack(fill="both", expand=True)
        else:
            self.html_renderer_widget = tk.Text(
                self.html_view_container,
                wrap="word",
                background="white",
                foreground="#222222",
                font=("Segoe UI", 11),
                padx=12,
                pady=12
            )
            self.html_renderer_widget.pack(fill="both", expand=True)

    def get_selected_preview_row_data(self):
        sample_row = {
            "Professor Name": "[Professor Name]",
            "Professor_Name": "[Professor Name]",
            "Institute Name": "[Institute Name]",
            "Institute_Name": "[Institute Name]",
            "Recent Paper Topics": "[Recent Paper Topics]",
            "Recent_Paper_Topics": "[Recent Paper Topics]",
            "Email Address": "[Email Address]",
            "Email_Address": "[Email Address]"
        }

        if self.recipients_df is not None and not self.recipients_df.empty:
            for col in self.recipients_df.columns:
                tag_label = f"[{col}]"
                sample_row[col] = tag_label
                sample_row[col.replace(" ", "_")] = tag_label

        selected_str = self.preview_target_menu.get()
        if selected_str == "Sample Placeholder Data" or self.recipients_df is None or self.recipients_df.empty:
            return sample_row

        if selected_str.startswith("Row "):
            try:
                idx = int(selected_str.split(":")[0].replace("Row ", "").strip()) - 1
                if 0 <= idx < len(self.recipients_df):
                    target_row = self.recipients_df.iloc[idx]
                    row_dict = {}
                    for col in self.recipients_df.columns:
                        val = str(target_row[col]) if pd.notna(target_row[col]) else ""
                        row_dict[col] = val
                        row_dict[col.replace(" ", "_")] = val
                    return row_dict
            except Exception:
                pass

        return sample_row

    def render_live_html_preview(self):
        if not hasattr(self, 'editor_text') or not self.editor_text.winfo_exists():
            return

        raw_template = self.editor_text.get("1.0", tk.END).strip()
        row_data = self.get_selected_preview_row_data()
        rendered_html = self.render_text_placeholders(raw_template, row_data)

        try:
            if HAS_TKINTERWEB and isinstance(self.html_renderer_widget, HtmlFrame):
                self.html_renderer_widget.load_html(rendered_html)
            elif HAS_TKHTMLVIEW and hasattr(self.html_renderer_widget, 'set_html'):
                self.html_renderer_widget.set_html(rendered_html)
            elif isinstance(self.html_renderer_widget, tk.Text):
                clean_text = re.sub(r'<[^>]+>', '', rendered_html)
                self.html_renderer_widget.config(state="normal")
                self.html_renderer_widget.delete("1.0", tk.END)
                self.html_renderer_widget.insert("1.0", clean_text)
                self.html_renderer_widget.config(state="disabled")
        except Exception as e:
            print(f"Live HTML render error: {e}")

    def update_preview_target_options(self):
        options = ["Sample Placeholder Data"]
        if self.recipients_df is not None and not self.recipients_df.empty:
            for idx, row in self.recipients_df.head(50).iterrows():
                p_name = row.get("Professor Name", row.get("Professor_Name", f"Contact {idx+1}"))
                inst = row.get("Institute Name", row.get("Institute_Name", ""))
                desc = f"Row {idx+1}: {p_name}"
                if inst:
                    desc += f" ({inst})"
                options.append(desc)

        prev_selection = self.preview_target_menu.get()
        self.preview_target_menu.configure(values=options)
        if prev_selection in options:
            self.preview_target_menu.set(prev_selection)
        elif len(options) > 1:
            self.preview_target_menu.set(options[1])
        else:
            self.preview_target_menu.set(options[0])

    def refresh_template_tag_buttons(self):
        if not hasattr(self, 'dynamic_tags_frame') or not self.dynamic_tags_frame.winfo_exists():
            return

        for child in self.dynamic_tags_frame.winfo_children():
            child.destroy()

        if self.recipients_df is not None and not self.recipients_df.empty:
            cols = [str(c) for c in self.recipients_df.columns]
        else:
            cols = ["Professor Name", "Institute Name", "Recent Paper Topics", "Email Address"]

        for col in cols:
            tag_name = f"{{{col}}}"
            btn = ctk.CTkButton(
                self.dynamic_tags_frame,
                text=tag_name,
                height=26,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#0f766e",
                hover_color="#0d9488",
                command=lambda t=tag_name: self.insert_placeholder(t)
            )
            btn.pack(side="left", padx=2, pady=2)

    def insert_placeholder(self, placeholder_text):
        self.editor_text.insert(tk.INSERT, placeholder_text)
        self.on_editor_key_release()

    def wrap_selected_html(self, start_tag, end_tag):
        try:
            sel_start = self.editor_text.index(tk.SEL_FIRST)
            sel_end = self.editor_text.index(tk.SEL_LAST)
            selected_text = self.editor_text.get(sel_start, sel_end)
            self.editor_text.delete(sel_start, sel_end)
            self.editor_text.insert(sel_start, f"{start_tag}{selected_text}{end_tag}")
        except tk.TclError:
            self.editor_text.insert(tk.INSERT, f"{start_tag}Text{end_tag}")
        self.on_editor_key_release()

    def show_insert_link_modal(self):
        modal = ctk.CTkToplevel(self)
        modal.title("Insert Hyperlink")
        modal.geometry("460x240")
        modal.minsize(400, 220)
        modal.grab_set()

        selected_text = ""
        try:
            sel_start = self.editor_text.index(tk.SEL_FIRST)
            sel_end = self.editor_text.index(tk.SEL_LAST)
            selected_text = self.editor_text.get(sel_start, sel_end)
        except tk.TclError:
            selected_text = "Click Here / Link"

        lbl1 = ctk.CTkLabel(modal, text="Text to Display:", font=ctk.CTkFont(weight="bold"))
        lbl1.pack(anchor="w", padx=20, pady=(15, 2))

        entry_text = ctk.CTkEntry(modal, placeholder_text="e.g. My Portfolio / Project Link")
        entry_text.pack(fill="x", padx=20, pady=(0, 10))
        entry_text.insert(0, selected_text)

        lbl2 = ctk.CTkLabel(modal, text="Web Address (URL):", font=ctk.CTkFont(weight="bold"))
        lbl2.pack(anchor="w", padx=20, pady=(0, 2))

        entry_url = ctk.CTkEntry(modal, placeholder_text="e.g. https://github.com/jahirmiru")
        entry_url.pack(fill="x", padx=20, pady=(0, 15))
        entry_url.insert(0, "https://")

        def _do_insert():
            display_txt = entry_text.get().strip() or "Link"
            url_target = entry_url.get().strip() or "https://"
            link_html = f'<a href="{url_target}" style="color: #0056b3;">{display_txt}</a>'
            
            try:
                sel_s = self.editor_text.index(tk.SEL_FIRST)
                sel_e = self.editor_text.index(tk.SEL_LAST)
                self.editor_text.delete(sel_s, sel_e)
                self.editor_text.insert(sel_s, link_html)
            except tk.TclError:
                self.editor_text.insert(tk.INSERT, link_html)

            modal.destroy()
            self.on_editor_key_release()

        btn_insert = ctk.CTkButton(
            modal, 
            text="Insert Link", 
            font=ctk.CTkFont(weight="bold"), 
            fg_color="#2563eb", 
            hover_color="#1d4ed8", 
            command=_do_insert
        )
        btn_insert.pack(pady=5)

    # ==========================================
    #     BACKGROUND SCHEDULER MONITOR ENGINE
    # ==========================================
    def start_background_scheduler_loop(self):
        def _loop():
            while self.scheduler_active:
                now_ts = time.time()
                for batch in self.scheduled_batches:
                    if batch.get("status") == "Pending":
                        try:
                            ts = batch.get("target_timestamp")
                            if not ts:
                                target_dt = datetime.datetime.strptime(batch["target_time"], "%Y-%m-%d %H:%M")
                                ts = target_dt.replace(tzinfo=BD_TIMEZONE).timestamp()
                            if now_ts >= ts:
                                self.after(0, lambda b=batch: self.execute_scheduled_batch(b))
                        except Exception as e:
                            print(f"Schedule check error: {e}")

                self.after(0, self.refresh_scheduled_batches_table)
                time.sleep(5)

        self.scheduler_worker_thread = threading.Thread(target=_loop, daemon=True)
        self.scheduler_worker_thread.start()

    def execute_scheduled_batch(self, batch):
        if self.is_sending:
            self.log_message(f"[!] Delaying batch {batch['id']} because campaign is currently busy...", "info")
            return

        batch["status"] = "Running"
        self.save_scheduled_queue()
        self.refresh_scheduled_batches_table()

        r_from = batch.get("row_from", 1)
        r_to = batch.get("row_to", len(self.recipients_df) if self.recipients_df is not None else 1)

        self.log_message(f"[+] Triggering Scheduled Batch {batch['id']} (Rows {r_from}-{r_to})", "success")
        self.start_campaign_subset(r_from, r_to, batch)

    # ==========================================
    #           DATA HANDLING & FILES
    # ==========================================
    def browse_recipients_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[
                ("Data Files (CSV, Excel)", "*.csv;*.xlsx;*.xls"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx;*.xls")
            ]
        )
        if filepath:
            self.load_recipients_from_file(filepath)

    def load_recipients_from_file(self, filepath):
        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)

            self.recipients_df = df
            self.data_filepath = filepath
            self.data_path_entry.delete(0, tk.END)
            self.data_path_entry.insert(0, filepath)

            self.populate_data_table(df)
            self.update_stats()
            self.update_preview_target_options()
            self.refresh_template_tag_buttons()
            self.trigger_live_preview_update()
            self.log_message(f"[+] Loaded {len(df)} rows from {os.path.basename(filepath)}", "success")

        except Exception as e:
            self.log_message(f"[-] Error loading data file: {e}", "error")
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def reload_current_recipients_file(self):
        filepath = self.data_path_entry.get().strip()
        if not filepath and self.data_filepath:
            filepath = self.data_filepath
        if filepath and os.path.exists(filepath):
            self.load_recipients_from_file(filepath)
            self.log_message(f"[✓] Data reloaded from: {os.path.basename(filepath)}", "success")
            messagebox.showinfo("Data Reloaded", f"Successfully reloaded {len(self.recipients_df)} rows from:\n{filepath}")
        else:
            messagebox.showwarning("No File to Reload", "No file loaded yet or file path is invalid. Please browse and select a file first.")

    def populate_data_table(self, df):
        self.data_tree.delete(*self.data_tree.get_children())

        if df is None or df.empty:
            self.data_tree["columns"] = []
            self.data_summary_label.configure(text="No data loaded. Supported columns: 'Professor Name', 'Institute Name', 'Email Address', 'Recent Paper Topics'")
            return

        cols = list(df.columns)
        tree_columns = ["#"] + cols
        self.data_tree["columns"] = tree_columns

        self.data_tree.heading("#", text="Row #")
        self.data_tree.column("#", width=65, minwidth=50, stretch=False, anchor="center")

        for col in cols:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=160, minwidth=100)

        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            vals = [str(idx)] + [str(val) if pd.notna(val) else "" for val in row]
            self.data_tree.insert("", tk.END, values=vals)

        email_col = self.find_email_column(df)
        if email_col:
            valid_emails = df[email_col].dropna().count()
            self.data_summary_label.configure(
                text=f"Total: {len(df)} rows | Columns: {len(cols)} | Valid Emails ({email_col}): {valid_emails}"
            )
        else:
            self.data_summary_label.configure(
                text=f"Total: {len(df)} rows | Warning: Could not auto-detect email column!"
            )

    def add_new_recipient(self):
        if self.recipients_df is None or self.recipients_df.empty:
            cols = ["Professor Name", "Institute Name", "Email Address", "Recent Paper Topics"]
            self.recipients_df = pd.DataFrame(columns=cols)
        else:
            cols = list(self.recipients_df.columns)

        modal = ctk.CTkToplevel(self)
        modal.title("Add New Recipient / Contact")
        modal.geometry("520x440")
        modal.minsize(450, 350)
        modal.grab_set()

        t_lbl = ctk.CTkLabel(modal, text="Enter Contact Information", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        t_lbl.pack(anchor="w", padx=20, pady=(15, 10))

        scroll_frame = ctk.CTkScrollableFrame(modal)
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        entries = {}
        for col in cols:
            lbl = ctk.CTkLabel(scroll_frame, text=f"{col}:", font=ctk.CTkFont(weight="bold"))
            lbl.pack(anchor="w", pady=(6, 2))
            ent = ctk.CTkEntry(scroll_frame, placeholder_text=f"Enter {col}...")
            ent.pack(fill="x", pady=(0, 4))
            entries[col] = ent

        def _do_add():
            new_row_data = {col: entries[col].get().strip() for col in cols}
            new_row_df = pd.DataFrame([new_row_data])
            self.recipients_df = pd.concat([self.recipients_df, new_row_df], ignore_index=True)
            self.populate_data_table(self.recipients_df)
            self.update_stats()
            self.update_preview_target_options()
            self.refresh_template_tag_buttons()
            self.trigger_live_preview_update()
            contact_label = new_row_data.get('Email Address', new_row_data.get('Professor Name', 'Contact'))
            self.log_message(f"[+] Added new recipient (Row #{len(self.recipients_df)}): {contact_label}", "success")
            modal.destroy()

        btn_bar = ctk.CTkFrame(modal, fg_color="transparent")
        btn_bar.pack(fill="x", padx=20, pady=12)

        btn_save = ctk.CTkButton(btn_bar, text="Save Contact", font=ctk.CTkFont(weight="bold"), fg_color="#0f766e", hover_color="#0d9488", command=_do_add)
        btn_save.pack(side="right", padx=5)

        btn_cancel = ctk.CTkButton(btn_bar, text="Cancel", fg_color="gray30", hover_color="gray40", command=modal.destroy)
        btn_cancel.pack(side="right", padx=5)

    def edit_selected_recipient(self):
        if self.recipients_df is None or self.recipients_df.empty:
            messagebox.showinfo("No Data", "No recipients loaded to edit.")
            return

        selected = self.data_tree.selection()
        if not selected:
            messagebox.showwarning("Select Row", "Please select a row from the table to edit.")
            return

        item = self.data_tree.item(selected[0])
        values = item["values"]
        if not values:
            return

        try:
            row_idx = int(values[0]) - 1
        except Exception:
            return

        if row_idx < 0 or row_idx >= len(self.recipients_df):
            return

        cols = list(self.recipients_df.columns)
        target_row = self.recipients_df.iloc[row_idx]

        modal = ctk.CTkToplevel(self)
        modal.title(f"Edit Recipient (Row #{row_idx + 1})")
        modal.geometry("520x440")
        modal.minsize(450, 350)
        modal.grab_set()

        t_lbl = ctk.CTkLabel(modal, text=f"Edit Information — Row #{row_idx + 1}", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        t_lbl.pack(anchor="w", padx=20, pady=(15, 10))

        scroll_frame = ctk.CTkScrollableFrame(modal)
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        entries = {}
        for col in cols:
            lbl = ctk.CTkLabel(scroll_frame, text=f"{col}:", font=ctk.CTkFont(weight="bold"))
            lbl.pack(anchor="w", pady=(6, 2))
            ent = ctk.CTkEntry(scroll_frame)
            val = str(target_row[col]) if pd.notna(target_row[col]) else ""
            ent.insert(0, val)
            ent.pack(fill="x", pady=(0, 4))
            entries[col] = ent

        def _do_save():
            for col in cols:
                self.recipients_df.at[row_idx, col] = entries[col].get().strip()
            
            self.populate_data_table(self.recipients_df)
            self.update_stats()
            self.update_preview_target_options()
            self.refresh_template_tag_buttons()
            self.trigger_live_preview_update()
            self.log_message(f"[✓] Updated data for Row #{row_idx + 1}", "success")
            modal.destroy()

        btn_bar = ctk.CTkFrame(modal, fg_color="transparent")
        btn_bar.pack(fill="x", padx=20, pady=12)

        btn_save = ctk.CTkButton(btn_bar, text="Update Changes", font=ctk.CTkFont(weight="bold"), fg_color="#2563eb", hover_color="#1d4ed8", command=_do_save)
        btn_save.pack(side="right", padx=5)

        btn_cancel = ctk.CTkButton(btn_bar, text="Cancel", fg_color="gray30", hover_color="gray40", command=modal.destroy)
        btn_cancel.pack(side="right", padx=5)

    def delete_selected_recipient(self):
        if self.recipients_df is None or self.recipients_df.empty:
            return

        selected = self.data_tree.selection()
        if not selected:
            messagebox.showwarning("Select Row", "Please select at least one row from the table to delete.")
            return

        indices_to_drop = []
        for item_id in selected:
            item = self.data_tree.item(item_id)
            values = item["values"]
            if values:
                try:
                    idx = int(values[0]) - 1
                    indices_to_drop.append(idx)
                except Exception:
                    pass

        if not indices_to_drop:
            return

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {len(indices_to_drop)} selected row(s)?"):
            self.recipients_df = self.recipients_df.drop(self.recipients_df.index[indices_to_drop]).reset_index(drop=True)
            self.populate_data_table(self.recipients_df)
            self.update_stats()
            self.update_preview_target_options()
            self.refresh_template_tag_buttons()
            self.trigger_live_preview_update()
            self.log_message(f"[!] Deleted {len(indices_to_drop)} row(s) from table.", "info")

    def save_current_recipients_file(self):
        if self.recipients_df is None or self.recipients_df.empty:
            messagebox.showinfo("No Data", "No recipient data to save.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV File", "*.csv"), ("Excel File", "*.xlsx")],
            initialfile=os.path.basename(self.data_filepath) if self.data_filepath else "recipients_list.csv"
        )
        if save_path:
            try:
                if save_path.endswith(".xlsx"):
                    self.recipients_df.to_excel(save_path, index=False)
                else:
                    self.recipients_df.to_csv(save_path, index=False)
                self.data_filepath = save_path
                self.data_path_entry.delete(0, tk.END)
                self.data_path_entry.insert(0, save_path)
                self.log_message(f"[✓] Saved updated recipients data to: {save_path}", "success")
                messagebox.showinfo("Saved", f"Recipient data successfully saved to:\n{save_path}")
            except Exception as e:
                self.log_message(f"[-] Failed to save recipient file: {e}", "error")
                messagebox.showerror("Save Failed", str(e))

    def find_email_column(self, df):
        if df is None:
            return None
        candidates = ["Email Address", "Email", "email", "E-mail", "EmailAddress", "Mail"]
        for c in candidates:
            if c in df.columns:
                return c
        for c in df.columns:
            if "email" in str(c).lower() or "mail" in str(c).lower():
                return c
        return None

    def deduplicate_recipients(self):
        if self.recipients_df is None or self.recipients_df.empty:
            messagebox.showwarning("No Data", "No data loaded to deduplicate.")
            return
        
        email_col = self.find_email_column(self.recipients_df)
        if not email_col:
            messagebox.showwarning("Error", "Could not identify email column.")
            return
        
        before = len(self.recipients_df)
        self.recipients_df.drop_duplicates(subset=[email_col], keep="first", inplace=True)
        self.recipients_df.reset_index(drop=True, inplace=True)
        after = len(self.recipients_df)
        
        self.populate_data_table(self.recipients_df)
        self.update_stats()
        self.update_preview_target_options()
        self.refresh_template_tag_buttons()
        self.trigger_live_preview_update()
        removed = before - after
        self.log_message(f"[+] Removed {removed} duplicate entries.", "info")
        messagebox.showinfo("Deduplicated", f"Removed {removed} duplicate email addresses.\nRemaining: {after} rows.")

    def remove_invalid_recipients(self):
        if self.recipients_df is None or self.recipients_df.empty:
            return
        email_col = self.find_email_column(self.recipients_df)
        if not email_col:
            return
        
        before = len(self.recipients_df)
        self.recipients_df = self.recipients_df[self.recipients_df[email_col].str.contains(r"^.+@.+\..+$", na=False)].reset_index(drop=True)
        after = len(self.recipients_df)
        
        self.populate_data_table(self.recipients_df)
        self.update_stats()
        self.update_preview_target_options()
        self.refresh_template_tag_buttons()
        self.trigger_live_preview_update()
        removed = before - after
        self.log_message(f"[+] Cleaned {removed} invalid/empty email rows.", "info")
        messagebox.showinfo("Cleaned", f"Removed {removed} invalid/empty rows.\nValid rows: {after}")

    # ==========================================
    #           ATTACHMENTS MANAGEMENT
    # ==========================================
    def browse_attachments(self):
        files = filedialog.askopenfilenames(
            title="Select Attachment Files",
            filetypes=[
                ("All Supported Files", "*.pdf;*.docx;*.doc;*.txt;*.png;*.jpg;*.zip"),
                ("PDF Documents", "*.pdf"),
                ("All Files", "*.*")
            ]
        )
        if files:
            for f in files:
                if f not in self.attachment_paths:
                    self.attachment_paths.append(f)
            self.update_attachments_display()

    def clear_attachments(self):
        self.attachment_paths = []
        self.update_attachments_display()

    def update_attachments_display(self):
        if not self.attachment_paths:
            self.attachments_label.configure(text="No files attached (Optional: PDF, DOCX, CV, Portfolio)", text_color="gray")
        else:
            names = [os.path.basename(p) for p in self.attachment_paths]
            self.attachments_label.configure(text=f"Attached ({len(names)}): {', '.join(names)}", text_color="#38bdf8")

    # ==========================================
    #           TEMPLATE MANAGEMENT
    # ==========================================
    def load_initial_defaults(self):
        default_sub = self.config_data.get("default_subject", "")
        if default_sub:
            self.subject_entry.insert(0, default_sub)

        if os.path.exists(DEFAULT_TEMPLATE_FILE):
            try:
                with open(DEFAULT_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                    content = f.read()
                    self.editor_text.insert("1.0", content)
            except Exception as e:
                print(f"Error loading template: {e}")

        if os.path.exists("sample_recipients.csv"):
            self.load_recipients_from_file("sample_recipients.csv")
        elif os.path.exists("Proff_List.csv"):
            self.load_recipients_from_file("Proff_List.csv")

        last_att = self.config_data.get("last_attachment", "")
        if last_att and os.path.exists(last_att):
            self.attachment_paths.append(os.path.abspath(last_att))
            self.update_attachments_display()

        self.apply_syntax_highlighting()
        self.trigger_live_preview_update()

    def load_template_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("HTML Template", "*.html;*.htm"), ("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                    self.editor_text.delete("1.0", tk.END)
                    self.editor_text.insert("1.0", content)
                self.apply_syntax_highlighting()
                self.trigger_live_preview_update()
                self.log_message(f"[+] Loaded template from {os.path.basename(path)}", "info")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load template:\n{e}")

    def save_template_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML Template", "*.html"), ("Text File", "*.txt")]
        )
        if path:
            try:
                content = self.editor_text.get("1.0", tk.END).strip()
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                messagebox.showinfo("Saved", "Template saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template:\n{e}")

    # ==========================================
    #         PREVIEW & TEST SEND
    # ==========================================
    def preview_sample_email(self):
        html_payload = self.editor_text.get("1.0", tk.END).strip()
        subject_raw = self.subject_entry.get().strip()

        if not html_payload:
            messagebox.showwarning("Warning", "Template is empty!")
            return

        sample_row = self.get_selected_preview_row_data()

        try:
            rendered_subj = self.render_text_placeholders(subject_raw, sample_row)
            rendered_body = self.render_text_placeholders(html_payload, sample_row)
        except Exception as e:
            messagebox.showerror("Template Error", f"Error rendering template placeholders:\n{e}")
            return

        preview_win = ctk.CTkToplevel(self)
        preview_win.title("Rendered Email Preview (Visual Output)")
        preview_win.geometry("840x670")
        preview_win.minsize(620, 460)
        preview_win.grab_set()

        header_frame = ctk.CTkFrame(preview_win)
        header_frame.pack(fill="x", padx=15, pady=(15, 10))

        sub_label = ctk.CTkLabel(header_frame, text=f"Subject: {rendered_subj}", font=ctk.CTkFont(size=13, weight="bold"), anchor="w")
        sub_label.pack(fill="x", padx=10, pady=(8, 4))

        recip_label = ctk.CTkLabel(
            header_frame, 
            text=f"Simulated Recipient: {sample_row.get('Professor_Name', 'Recipient')} ({sample_row.get('Email_Address', 'sample@example.com')}) @ {sample_row.get('Institute_Name', 'Organization')}", 
            font=ctk.CTkFont(size=11), 
            text_color="gray", 
            anchor="w"
        )
        recip_label.pack(fill="x", padx=10, pady=(0, 8))

        html_container = ctk.CTkFrame(preview_win, fg_color="white", corner_radius=6)
        html_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        if HAS_TKINTERWEB:
            modal_html = HtmlFrame(html_container, horizontal_scrollbar="auto", vertical_scrollbar="auto")
            modal_html.pack(fill="both", expand=True)
            modal_html.load_html(rendered_body)
        elif HAS_TKHTMLVIEW:
            modal_html = tkhtmlview.HTMLText(html_container, background="white", foreground="black", font=("Arial", 11))
            modal_html.pack(fill="both", expand=True)
            modal_html.set_html(rendered_body)
        else:
            modal_text = tk.Text(html_container, wrap="word", background="white", foreground="#222222", font=("Segoe UI", 11), padx=12, pady=12)
            modal_text.pack(fill="both", expand=True)
            modal_text.insert("1.0", re.sub(r'<[^>]+>', '', rendered_body))
            modal_text.config(state="disabled")

    def show_test_email_dialog(self):
        dialog = ctk.CTkInputDialog(
            text="Enter destination email address for test send:\n(A sample email with attachments will be delivered)",
            title="Send Test Email"
        )
        target_email = dialog.get_input()

        if target_email and "@" in target_email:
            self.send_test_email_action(target_email.strip())

    def send_test_email_action(self, target_email):
        def _task():
            self.log_message(f"[...] Sending test email to {target_email}...", "info")
            try:
                my_email = self.smtp_email_entry.get().strip()
                my_password = self.smtp_pass_entry.get().strip()
                host = self.smtp_host_entry.get().strip()
                port = int(self.smtp_port_entry.get().strip() or 587)
                use_tls = self.smtp_tls_var.get()
                sender_name = self.sender_name_entry.get().strip()
                html_raw = self.editor_text.get("1.0", tk.END).strip()
                subject_raw = self.subject_entry.get().strip()
                sample_data = {
                    "Professor Name": "Test Recipient",
                    "Professor_Name": "Test Recipient",
                    "Institute Name": "Test Research Lab",
                    "Institute_Name": "Test Research Lab",
                    "Recent Paper Topics": "Traffic Microsimulation & AI Calibration",
                    "Recent_Paper_Topics": "Traffic Microsimulation & AI Calibration",
                    "Email Address": target_email,
                    "Email_Address": target_email
                }

                subj = self.render_text_placeholders(subject_raw, sample_data)
                body = self.render_text_placeholders(html_raw, sample_data)

                msg = MIMEMultipart()
                if sender_name:
                    msg['From'] = f"{sender_name} <{my_email}>"
                else:
                    msg['From'] = my_email
                msg['To'] = target_email
                msg['Subject'] = f"[TEST] {subj}"
                msg['Date'] = formatdate(localtime=True)
                msg.attach(MIMEText(body, 'html'))

                for att_path in self.attachment_paths:
                    if os.path.exists(att_path):
                        with open(att_path, "rb") as f:
                            part = MIMEApplication(f.read(), _subtype="octet-stream")
                            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(att_path))
                            msg.attach(part)

                server = smtplib.SMTP(host, port, timeout=20)
                if use_tls:
                    server.starttls()
                server.login(my_email, my_password)
                server.send_message(msg)
                server.quit()

                self.log_message(f"[✓] Test email successfully delivered to {target_email}!", "success")
                messagebox.showinfo("Success", f"Test email sent successfully to {target_email}!")
            except Exception as e:
                self.log_message(f"[X] Failed to send test email: {e}", "error")
                messagebox.showerror("Error", f"Failed to send test email:\n{e}")

        threading.Thread(target=_task, daemon=True).start()

    # ==========================================
    #         PLACEHOLDER REPLACEMENT
    # ==========================================
    def render_text_placeholders(self, template_str, row_dict):
        result = template_str
        for key, val in row_dict.items():
            str_val = str(val) if pd.notna(val) else ""
            result = result.replace(f"{{{key}}}", str_val)
            result = result.replace(f"{{{key.replace(' ', '_')}}}", str_val)
            result = result.replace(f"{{{key.replace('_', ' ')}}}", str_val)
        return result

    # ==========================================
    #         CAMPAIGN DISPATCH ENGINE
    # ==========================================
    def update_stats(self):
        if self.recipients_df is not None:
            self.total_targets = len(self.recipients_df)
        else:
            self.total_targets = 0

        self.card_total.configure(text=str(self.total_targets))
        self.card_sent.configure(text=str(self.sent_count))
        self.card_failed.configure(text=str(self.failed_count))

    def start_campaign(self):
        if self.recipients_df is None or self.recipients_df.empty:
            messagebox.showwarning("No Data", "Please load a CSV or Excel recipient list first.")
            return
        
        confirm = messagebox.askyesno(
            "Start Immediate Campaign", 
            f"Are you ready to send automated emails to {len(self.recipients_df)} recipients?"
        )
        if confirm:
            self.start_campaign_subset(1, len(self.recipients_df))

    def start_campaign_subset(self, start_row, end_row, linked_batch=None):
        if self.is_sending:
            return

        my_email = self.smtp_email_entry.get().strip()
        my_password = self.smtp_pass_entry.get().strip()

        if not my_email or not my_password:
            messagebox.showwarning("Missing Credentials", "Please enter your SMTP Email and App Password in Settings.")
            self.show_page("settings")
            return

        html_content = self.editor_text.get("1.0", tk.END).strip()
        if not html_content:
            messagebox.showwarning("Empty Template", "Please write an email template in Template Editor.")
            self.show_page("template")
            return

        email_col = self.find_email_column(self.recipients_df)
        if not email_col:
            messagebox.showerror("Error", "Could not find Email Address column in data file.")
            return

        self.is_sending = True
        self.is_paused = False
        self.should_stop = False
        self.sent_count = 0
        self.failed_count = 0
        self.skipped_count = 0

        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal", text="Pause")
        self.btn_stop.configure(state="normal")
        self.card_status.configure(text="RUNNING", text_color="#10b981")

        self.save_config()

        self.campaign_thread = threading.Thread(
            target=lambda: self.campaign_worker_subset(start_row, end_row, linked_batch),
            daemon=True
        )
        self.campaign_thread.start()

    def toggle_pause(self):
        if not self.is_sending:
            return
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.btn_pause.configure(text="Resume", fg_color="#10b981", hover_color="#059669")
            self.card_status.configure(text="PAUSED", text_color="#f59e0b")
            self.log_message("[||] Campaign paused.", "info")
        else:
            self.btn_pause.configure(text="Pause", fg_color="#f59e0b", hover_color="#d97706")
            self.card_status.configure(text="RUNNING", text_color="#10b981")
            self.log_message("[▶] Campaign resumed.", "info")

    def stop_campaign(self):
        if not self.is_sending:
            return
        self.should_stop = True
        self.log_message("[⏹] Stopping campaign after current item...", "error")

    def campaign_worker_subset(self, start_row, end_row, linked_batch=None):
        my_email = self.smtp_email_entry.get().strip()
        my_password = self.smtp_pass_entry.get().strip()
        host = self.smtp_host_entry.get().strip()
        port = int(self.smtp_port_entry.get().strip() or 587)
        use_tls = self.smtp_tls_var.get()
        sender_name = self.sender_name_entry.get().strip()

        subject_raw = self.subject_entry.get().strip()
        html_template = self.editor_text.get("1.0", tk.END).strip()

        base_delay = int(self.delay_spinbox.get() or 12)
        jitter = int(self.jitter_spinbox.get() or 4)
        batch_pause_count = int(self.batch_count_spinbox.get() or 20)
        batch_pause_mins = int(self.batch_pause_spinbox.get() or 5)

        full_df = self.recipients_df
        s_idx = max(0, start_row - 1)
        e_idx = min(len(full_df), end_row)
        subset_df = full_df.iloc[s_idx:e_idx]
        total_rows = len(subset_df)
        email_col = self.find_email_column(full_df)

        self.log_message(f"=== Starting Campaign for {total_rows} recipients (Rows {start_row}-{end_row}) ===", "info")

        server = None
        try:
            self.log_message(f"[+] Connecting to {host}:{port}...", "info")
            server = smtplib.SMTP(host, port, timeout=25)
            if use_tls:
                server.starttls()
            server.login(my_email, my_password)
            self.log_message("[+] SMTP Server authenticated successfully!", "success")
        except Exception as e:
            self.log_message(f"[-] SMTP Connection Error: {e}", "error")
            if linked_batch:
                linked_batch["status"] = "Failed"
                self.save_scheduled_queue()
            self.finish_campaign("FAILED")
            return

        emails_sent_in_batch = 0

        for current_idx, (actual_row_idx, row) in enumerate(subset_df.iterrows(), start=1):
            if self.should_stop:
                self.log_message("[!] Campaign stopped by user.", "error")
                break

            while self.is_paused and not self.should_stop:
                time.sleep(0.5)

            if self.should_stop:
                break

            row_dict = {}
            for col in full_df.columns:
                row_dict[col] = str(row[col]) if pd.notna(row[col]) else ""
                row_dict[col.replace(" ", "_")] = str(row[col]) if pd.notna(row[col]) else ""

            recipient_email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
            prof_name = row_dict.get("Professor Name", row_dict.get("Professor_Name", "Recipient"))
            institute = row_dict.get("Institute Name", row_dict.get("Institute_Name", "Organization"))

            if not recipient_email or "@" not in recipient_email or recipient_email.lower() == "nan":
                self.skipped_count += 1
                self.log_message(f"[-] Row {actual_row_idx+1}: Skipped (Missing/Invalid Email)", "error")
                continue

            try:
                filled_subject = self.render_text_placeholders(subject_raw, row_dict)
                filled_body = self.render_text_placeholders(html_template, row_dict)

                msg = MIMEMultipart()
                if sender_name:
                    msg['From'] = f"{sender_name} <{my_email}>"
                else:
                    msg['From'] = my_email
                msg['To'] = recipient_email
                msg['Subject'] = filled_subject
                msg['Date'] = formatdate(localtime=True)
                msg.attach(MIMEText(filled_body, 'html'))

                for att_path in self.attachment_paths:
                    if os.path.exists(att_path):
                        with open(att_path, "rb") as f:
                            part = MIMEApplication(f.read(), _subtype="octet-stream")
                            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(att_path))
                            msg.attach(part)

                server.send_message(msg)
                self.sent_count += 1
                emails_sent_in_batch += 1
                self.log_message(f"[✓] [{current_idx}/{total_rows}] Sent to {prof_name} ({recipient_email}) @ {institute}", "success")
                
                self.campaign_log_records.append({
                    "Index": actual_row_idx + 1,
                    "Name": prof_name,
                    "Email": recipient_email,
                    "Institute": institute,
                    "Status": "Sent",
                    "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Error": ""
                })

            except smtplib.SMTPServerDisconnected:
                self.log_message(f"[!] Server disconnected. Reconnecting for row {actual_row_idx+1}...", "error")
                try:
                    server = smtplib.SMTP(host, port, timeout=25)
                    if use_tls:
                        server.starttls()
                    server.login(my_email, my_password)
                    server.send_message(msg)
                    self.sent_count += 1
                    emails_sent_in_batch += 1
                    self.log_message(f"[✓] [{current_idx}/{total_rows}] Sent after reconnect to {prof_name}", "success")
                except Exception as ex2:
                    self.failed_count += 1
                    self.log_message(f"[X] Failed sending to {prof_name}: {ex2}", "error")
            except Exception as e:
                self.failed_count += 1
                self.log_message(f"[X] [{current_idx}/{total_rows}] FAILED sending to {prof_name} ({recipient_email}): {e}", "error")

            current_progress = current_idx / total_rows
            self.after(0, lambda p=current_progress, i=current_idx, t=total_rows: self.update_progress_ui(p, i, t))

            # Batch Cooldown Pause
            if batch_pause_count > 0 and emails_sent_in_batch % batch_pause_count == 0 and current_idx < total_rows and not self.should_stop:
                self.log_message(f"[☕] Batch limit of {batch_pause_count} reached. Cooldown pause for {batch_pause_mins} minutes...", "info")
                for _ in range(batch_pause_mins * 60):
                    if self.should_stop:
                        break
                    time.sleep(1)

            # Normal Delay / Jitter Sleep
            if current_idx < total_rows and not self.should_stop:
                import random
                sleep_time = max(1, base_delay + random.randint(-jitter, jitter) if jitter > 0 else base_delay)
                self.log_message(f"[⏳] Waiting {sleep_time}s before next email...", "info")
                for _ in range(int(sleep_time * 2)):
                    if self.should_stop:
                        break
                    time.sleep(0.5)

        try:
            if server:
                server.quit()
        except Exception:
            pass

        if linked_batch:
            linked_batch["status"] = "Completed" if not self.should_stop else "Stopped"
            self.save_scheduled_queue()
            self.refresh_scheduled_batches_table()

        status_str = "STOPPED" if self.should_stop else "COMPLETED"
        self.finish_campaign(status_str)

    def update_progress_ui(self, fraction, current, total):
        self.progress_bar.set(fraction)
        self.progress_label.configure(
            text=f"Progress: {current}/{total} ({int(fraction * 100)}%) — Sent: {self.sent_count} | Failed: {self.failed_count} | Skipped: {self.skipped_count}"
        )
        self.update_stats()

    def finish_campaign(self, status):
        def _cleanup():
            self.is_sending = False
            self.btn_start.configure(state="normal")
            self.btn_pause.configure(state="disabled", text="Pause")
            self.btn_stop.configure(state="disabled")
            
            if status == "COMPLETED":
                self.card_status.configure(text="DONE", text_color="#10b981")
                self.log_message(f"=== Campaign Finished: {self.sent_count} sent, {self.failed_count} failed ===", "success")
                messagebox.showinfo("Campaign Completed", f"Campaign completed!\n\nSent: {self.sent_count}\nFailed: {self.failed_count}\nSkipped: {self.skipped_count}")
            else:
                self.card_status.configure(text=status, text_color="#ef4444")
                self.log_message(f"=== Campaign ended with status: {status} ===", "error")
            
            self.update_stats()

        self.after(0, _cleanup)

# ==========================================
#               ENTRY POINT
# ==========================================
if __name__ == "__main__":
    app = AutoMailerApp()
    app.mainloop()
