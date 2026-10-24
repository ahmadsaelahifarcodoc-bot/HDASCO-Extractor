import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

from chrome_cdp import ChromeCDP
from extractors import extract_br_list, download_bl_xmls, export_xml_folder, extract_container_history

APP_NAME = "HDASCO Extractor"
HOME = Path.home()
EXPORT_DIR = HOME / "Documents" / "HDASCO_Exports"
PROFILE_DIR = HOME / "AppData" / "Local" / "HDASCOExtractor" / "ChromeProfile"

BR_URL = "https://my.hdasco.com/Packages/Documentation/BRList.aspx"
BL_URL = "https://my.hdasco.com/Packages/Documentation/BLList.aspx"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("760x510")
        self.minsize(700, 470)
        self.cdp = None
        self.busy = False
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)

        ttk.Label(self, text="HDASCO Extractor", font=("Segoe UI", 18, "bold")).pack(pady=(18, 4))
        ttk.Label(self, text="Python + Google Chrome (dedicated profile; no Selenium)",
                  font=("Segoe UI", 10)).pack(pady=(0, 16))

        browser = ttk.LabelFrame(self, text="1. Chrome / HDASCO")
        browser.pack(fill="x", padx=18, pady=6)
        ttk.Button(browser, text="Open HDASCO in dedicated Chrome", command=self.open_chrome).pack(
            side="left", padx=10, pady=10)
        ttk.Button(browser, text="Connect to Chrome", command=self.connect_chrome).pack(
            side="left", padx=5, pady=10)
        ttk.Button(browser, text="Open BR List", command=lambda: self.navigate(BR_URL)).pack(
            side="left", padx=5, pady=10)
        ttk.Button(browser, text="Open BL List", command=lambda: self.navigate(BL_URL)).pack(
            side="left", padx=5, pady=10)

        opts = ttk.LabelFrame(self, text="2. Extraction options")
        opts.pack(fill="x", padx=18, pady=6)
        ttk.Label(opts, text="Pages to process:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.pages = tk.StringVar(value="2")
        ttk.Spinbox(opts, from_=1, to=500, width=8, textvariable=self.pages).grid(
            row=0, column=1, padx=4, pady=10, sticky="w")
        ttk.Label(opts, text="Delay between pages (seconds):").grid(
            row=0, column=2, padx=(20, 4), pady=10, sticky="w")
        self.delay = tk.StringVar(value="2")
        ttk.Spinbox(opts, from_=0, to=60, increment=1, width=8, textvariable=self.delay).grid(
            row=0, column=3, padx=4, pady=10, sticky="w")

        actions = ttk.LabelFrame(self, text="3. Extract / export")
        actions.pack(fill="x", padx=18, pady=6)
        ttk.Button(actions, text="Extract BR List to Excel", command=lambda: self.run_task(
            "BR extraction", lambda: extract_br_list(self.get_cdp(), self.page_count(), self.delay_seconds(), EXPORT_DIR, self.log))).grid(
            row=0, column=0, padx=8, pady=9, sticky="ew")
        ttk.Button(actions, text="Download BL XML files", command=lambda: self.run_task(
            "BL XML download", lambda: download_bl_xmls(self.get_cdp(), self.page_count(), self.delay_seconds(), EXPORT_DIR, self.log))).grid(
            row=0, column=1, padx=8, pady=9, sticky="ew")
        ttk.Button(actions, text="Export downloaded XML to Excel", command=lambda: self.run_task(
            "XML to Excel", lambda: export_xml_folder(EXPORT_DIR / "XML", EXPORT_DIR, self.log))).grid(
            row=1, column=0, padx=8, pady=9, sticky="ew")
        ttk.Button(actions, text="Extract container history", command=lambda: self.run_task(
            "Container history", lambda: extract_container_history(self.get_cdp(), self.page_count(), self.delay_seconds(), EXPORT_DIR, self.log))).grid(
            row=1, column=1, padx=8, pady=9, sticky="ew")
        for c in range(2):
            actions.columnconfigure(c, weight=1)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=18, pady=(6, 0))
        ttk.Button(bottom, text="Open export folder", command=self.open_export_folder).pack(side="left")
        ttk.Button(bottom, text="Disconnect", command=self.disconnect).pack(side="left", padx=8)
        self.status = tk.StringVar(value="Ready. Open Chrome, sign in to HDASCO, then connect.")
        ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken").pack(
            fill="x", side="bottom", padx=8, pady=8)

    def log(self, message):
        self.after(0, lambda: self.status.set(str(message)[:220]))

    def page_count(self):
        try:
            n = int(self.pages.get())
            if n < 1:
                raise ValueError()
            return n
        except ValueError:
            raise ValueError("Pages to process must be a whole number from 1 to 500.")

    def delay_seconds(self):
        try:
            return max(0, min(60, float(self.delay.get())))
        except ValueError:
            raise ValueError("Delay must be a number from 0 to 60 seconds.")

    def open_chrome(self):
        try:
            from chrome_cdp import launch_chrome
            launch_chrome(PROFILE_DIR)
            self.log("Chrome started. Sign in to HDASCO, then click Connect to Chrome.")
        except Exception as e:
            messagebox.showerror(APP_NAME, str(e))

    def connect_chrome(self):
        try:
            self.cdp = ChromeCDP.connect()
            self.log("Connected to Chrome. Current tab: " + self.cdp.current_url())
        except Exception as e:
            messagebox.showerror(APP_NAME, "Could not connect to Chrome.\n\n"
                "1. Click Open HDASCO in dedicated Chrome.\n"
                "2. Sign in if needed.\n"
                "3. Click Connect to Chrome.\n\nDetails: " + str(e))

    def get_cdp(self):
        if not self.cdp:
            self.cdp = ChromeCDP.connect()
        return self.cdp

    def navigate(self, url):
        try:
            self.get_cdp().navigate(url)
            self.log("Opened " + url + " — sign in if the site asks.")
        except Exception as e:
            messagebox.showerror(APP_NAME, str(e))

    def run_task(self, name, task):
        if self.busy:
            messagebox.showinfo(APP_NAME, "Another task is already running.")
            return
        self.busy = True
        self.log(name + " started...")
        def worker():
            try:
                result = task()
                self.log(name + " finished. " + (str(result) if result else ""))
                self.after(0, lambda: messagebox.showinfo(APP_NAME, name + " finished.\n" + (str(result) if result else "")))
            except Exception as e:
                self.log(name + " failed: " + str(e))
                self.after(0, lambda err=str(e): messagebox.showerror(APP_NAME, name + " failed:\n\n" + err))
            finally:
                self.busy = False
        threading.Thread(target=worker, daemon=True).start()

    def open_export_folder(self):
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(str(EXPORT_DIR))

    def disconnect(self):
        self.cdp = None
        self.log("Disconnected from Chrome. Chrome remains open.")


if __name__ == "__main__":
    App().mainloop()
